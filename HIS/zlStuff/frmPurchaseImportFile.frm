VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPurchaseImportFile 
   Caption         =   "�����ⲿ�ļ�"
   ClientHeight    =   6630
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   9075
   Icon            =   "frmPurchaseImportFile.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   9075
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ImageList imgError 
      Left            =   5400
      Top             =   840
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
            Picture         =   "frmPurchaseImportFile.frx":6852
            Key             =   "error"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseImportFile.frx":D0B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplit 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   240
      MousePointer    =   7  'Size N S
      ScaleHeight     =   300
      ScaleWidth      =   4215
      TabIndex        =   10
      Top             =   4560
      Width           =   4215
      Begin VB.Label lblCollect 
         AutoSize        =   -1  'True
         Caption         =   "�ɱ���     Ԫ          ��Ʊ��     Ԫ"
         Height          =   180
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   3960
      End
   End
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   3360
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.xls|*.xls|*.xlsx|*.xlsx"
   End
   Begin MSComctlLib.ProgressBar ProCheck 
      Height          =   300
      Left            =   360
      TabIndex        =   7
      Top             =   6000
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdProvider 
      Caption         =   "��"
      Height          =   300
      Left            =   4440
      TabIndex        =   5
      Top             =   960
      Width           =   280
   End
   Begin VB.TextBox txtProvider 
      Height          =   300
      Left            =   960
      TabIndex        =   4
      Top             =   960
      Width           =   3495
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "��"
      Height          =   300
      Left            =   4440
      TabIndex        =   2
      Top             =   480
      Width           =   280
   End
   Begin VB.TextBox txtFile 
      Enabled         =   0   'False
      Height          =   300
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.ComboBox cboIOType 
      Height          =   300
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   960
      Width           =   3780
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   2565
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   4935
      _cx             =   8705
      _cy             =   4524
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
      BackColorSel    =   4227072
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
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
      FormatString    =   $"frmPurchaseImportFile.frx":13916
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
      ExplorerBar     =   1
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
      VirtualData     =   0   'False
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
   Begin VSFlex8Ctl.VSFlexGrid vsfError 
      Height          =   765
      Left            =   240
      TabIndex        =   9
      Top             =   5040
      Width           =   4935
      _cx             =   8705
      _cy             =   1349
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
      BackColorSel    =   12632256
      ForeColorSel    =   16711680
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
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
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPurchaseImportFile.frx":1398B
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
      ExplorerBar     =   1
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
      VirtualData     =   0   'False
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
   Begin XtremeCommandBars.ImageManager imgPicture 
      Left            =   840
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmPurchaseImportFile.frx":13A00
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label lblProvider 
      AutoSize        =   -1  'True
      Caption         =   "��Ӧ��(&P)"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   1005
      Width           =   810
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      Caption         =   "��  ��(&F)"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   525
      Width           =   810
   End
End
Attribute VB_Name = "frmPurchaseImportFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MCONTOOLMODE As Integer = 100 '����
Private Const MCONTOOLOUTPUT As Integer = 101   '�����ļ�
Private Const MCONTOOLCHECK As Integer = 102    'У��
Private Const MCONTOOLSAVE As Integer = 103 '����
Private Const MCONTOOLEXIT As Integer = 104 '�˳�
Private Const MCONERROR As Integer = 105 '������ʾͼ��
Private Const MCONWARN   As Integer = 106   '����ͼ��
Private Const MCONYESCHECK As Integer = 107 '�ϸ���
Private Const mconNOCHECK As Integer = 108 '���ϸ���
Private Const MCONTOOLCHECKCONDITION As Integer = 107 '�������

Private Mstr_Cols As String  'EXCEL��������
Private mblnResult As Boolean
Private mblnChange As Boolean
Private mlngModule As Long
Private mlngStockID As Long
Private mstrStock As String
Private mblnVirtualStock As Boolean '�Ƿ�������ⷿ��true-�� false-����
Private mintUnit  As Integer                    '��ʾ��λ:0-ɢװ��λ,1-��װ��λ
Private mFMT As g_FmtString
'����Ϊ���뷽ʽ/���ı���|����|�ɱ���|�ɱ����|��Ʊ���|����*�ɱ���=�ɱ����|��Ʊ���=�ɱ����|���ɱ���=HIS�ɱ���|Ч��|�������|���Ч��|��������|�洢�ⷿ|����ⷿ|��Ʒ����(0-����ȫ����1-��ȫ����/0-��ʾ1-��ֹ|....)
Private mbyt���뷽ʽ, mbyt���ı���, mbyt����, mbyt�ɱ���, mbyt�ɱ����, mbyt��Ʊ���, mbyt��Ʊ����, mbytNumCost, mbytInvoiceCost, mbytExcelCost, mbytЧ��, mbyt�������, mbyt���Ч��, mbyt��������, mbyt�洢�ⷿ, mbyt����ⷿ, mbyt��Ʒ���� As Byte

'Private mobjXLS As Excel.Application
'Private mobjWB As Excel.Workbook
'Private mobjWS As Excel.Worksheet
Private mobjXLS As Object
Private mobjWB As Object
Private mobjWS As Object

Private Sub InitComandbar()
    '��ʼ��������
    Dim cbrControlMain As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim cbrToolPopup As CommandBarPopup
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsMain.VisualTheme = xtpThemeOffice2003

    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsMain.EnableCustomization False
    Me.cbsMain.Icons = Me.imgPicture.Icons
    
    '����������
    Set cbrToolBar = Me.cbsMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched Or xtpFlagFloating Or xtpFlagAlignAny
        
    With cbrToolBar.Controls
'        Set cbrControlMain = .Add(xtpControlSplitButtonPopup, MCONTOOLMODE, "Excel����")
'        cbrControlMain.CommandBar.Controls.Add xtpControlButton, MCONTOOLMODEEXCEL, "Excel����"    '�����Ӳ˵�
'        cbrControlMain.CommandBar.Controls.Add xtpControlButton, MCONTOOLMODEXML, "XML����"
'        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        
'        Set cbrControlMain = .Add(xtpControlSplitButtonPopup, MCONTOOLOUTPUT, "����Excel")
'        cbrControlMain.CommandBar.Controls.Add xtpControlButton, MCONTOOLOUTPUTEXCEL, "����Excel"  '�����ļ��Ӳ˵�
'        cbrControlMain.CommandBar.Controls.Add xtpControlButton, MCONTOOLOUTPUTXML, "����XML"
        Set cbrControlMain = .Add(xtpControlButton, MCONTOOLMODE, "Excel����")
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        Set cbrControlMain = .Add(xtpControlButton, MCONTOOLOUTPUT, "����Excel")
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        Set cbrControlMain = .Add(xtpControlButton, MCONTOOLCHECKCONDITION, "�������")
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, MCONTOOLCHECK, "У��")
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, MCONTOOLSAVE, "����")
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        Set cbrControlMain = .Add(xtpControlButton, MCONTOOLEXIT, "�˳�")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
    End With
    cbsMain.Item(1).Delete
End Sub

Public Property Get Result() As Boolean
    Result = mblnResult
End Property

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case MCONTOOLMODE    '��������
            Call ProduceStyleBook
        Case MCONTOOLOUTPUT '�����ļ�
            Call OutPutFile
        Case MCONTOOLCHECK  '����У��
            Call CheckData
        Case MCONTOOLSAVE   '����
            Call SaveCard
        Case MCONTOOLCHECKCONDITION '��������
            frmPurchaseImportFileCondition.ShowMe Me, mlngModule
            If vsfList.Rows > 1 Then Call CheckData
        Case MCONTOOLEXIT    '�˳�
            Unload Me
    End Select
End Sub

Private Sub OutPutFile()
    '��������ļ�
    Dim strFileName As String
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lng���� As Long
    
    On Error GoTo ErrHandle
    Call InitExcel
    Set mobjWB = mobjXLS.Workbooks.Add
    Set mobjWS = mobjWB.ActiveSheet
    
    With vsfList
        If .Rows = 1 Then Exit Sub
        lng���� = GetColumnPostation("���ı���") + Asc("A") - 1 '�ߵ���һ��֤���϶������ı�����
        mobjWS.Range(Chr(lng����) & "1:" & Chr(lng����) & .Rows - 1).NumberFormatLocal = "@"
        For lngRow = 0 To .Rows - 1
            For lngCol = 1 To .Cols - 1
                If lngCol = 1 And lngRow <> 0 Then
                    mobjWS.cells(lngRow + 1, lngCol) = "'" & .TextMatrix(lngRow, lngCol)
                Else
                    mobjWS.cells(lngRow + 1, lngCol) = .TextMatrix(lngRow, lngCol)
                End If
            Next
        Next
    End With
    
    With dlgOpenFile
        .CancelError = True
        .FileName = ""
        .Filter = "*.xls|*.xls|*.xlsx|*.xlsx"
        .ShowSave
        strFileName = .FileName
        If Trim(strFileName) <> "" Then
            mobjWB.SaveAs strFileName
            mobjWB.Close
            Set mobjWS = Nothing
            Set mobjWB = Nothing
            mobjXLS.quit
            MsgBox "����ɹ���", vbInformation, gstrSysName
        End If
    End With
    Exit Sub
    
ErrHandle:
    mobjWB.Close
    Set mobjWS = Nothing
    Set mobjWB = Nothing
    mobjXLS.quit
End Sub

Private Sub chkNoCheck_Click()
    If vsfList.Rows > 1 Then
        Call CheckData
    End If
End Sub

Private Sub chkYesCheck_Click()
    If vsfList.Rows > 1 Then
        Call CheckData
    End If
End Sub

Private Sub cmdFile_Click()
    On Error GoTo ErrHandle
    
    dlgOpenFile.FileName = ""
    dlgOpenFile.Filter = "*.xls|*.xls|*.xlsx|*.xlsx"
    dlgOpenFile.ShowOpen
    If dlgOpenFile.FileName <> "" Then
        txtFile.Text = dlgOpenFile.FileName
        If mlngModule = 1712 Then
            txtProvider.SetFocus
        ElseIf mlngModule = 1714 Then
            cboIOType.SetFocus
        End If
    Else
        GoTo ErrHandle
    End If
    If txtFile.Text <> "" Then
        Call ParseParameter
        DoEvents
        Call FS.ShowFlash("���ڼ�������,���Ժ� ...", Me)
        Me.MousePointer = vbHourglass
        
        ProCheck.Value = 0
        ProCheck.Visible = True
        Call InitExcel
        Call GetExcelData
        
        Me.MousePointer = vbDefault
        Call FS.StopFlash
        ProCheck.Visible = False
    End If
    Exit Sub
    
ErrHandle:
    Exit Sub
End Sub

Private Sub ParseParameter()
    '��������
    Dim i As Integer
    Dim arryPara As Variant
    Dim strPara As String
    
    If mlngModule = 1712 Then
        strPara = zlDatabase.GetPara("�����ļ���鷽ʽ", glngSys, mlngModule, "0/0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0")
    Else
        strPara = zlDatabase.GetPara("�����ļ���鷽ʽ", glngSys, mlngModule, "0/0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0")
    End If
    
    mbyt���뷽ʽ = Mid(strPara, 1, 1)
    strPara = Mid(strPara, 3)
    arryPara = Split(strPara, "|")
    mbyt���ı��� = arryPara(0)
    mbyt���� = arryPara(1)
    mbyt�ɱ��� = arryPara(2)
    mbyt�ɱ���� = arryPara(3)
    mbyt��Ʊ��� = arryPara(4)
    mbyt��Ʊ���� = arryPara(5)
    mbytNumCost = arryPara(6)
    mbytInvoiceCost = arryPara(7)
    mbytExcelCost = arryPara(8)
    mbytЧ�� = arryPara(9)
    mbyt������� = arryPara(10)
    mbyt���Ч�� = arryPara(11)
    mbyt�������� = arryPara(12)
    mbyt�洢�ⷿ = arryPara(13)
    mbyt����ⷿ = arryPara(14)
    mbyt��Ʒ���� = arryPara(15)
End Sub

Private Sub ProduceStyleBook()
'���ɵ����ⲿ�ļ��ı�׼XLS�ļ�����
    Dim arrCols As Variant
    Dim i As Byte
    Dim blnFinished As Boolean
    Dim strFileName As String
    
    arrCols = Split(Mstr_Cols, ";")
    
    On Error GoTo ErrHandle
    Call InitExcel
    Set mobjWB = mobjXLS.Workbooks.Add
    Set mobjWS = mobjWB.ActiveSheet
    
    For i = LBound(arrCols) + 1 To UBound(arrCols)
        mobjWS.cells(1, i) = arrCols(i)
    Next
    
    With dlgOpenFile
        .FileName = ""
        .Filter = "Excel Files (*.xls)|*.xls"
        .ShowSave
        strFileName = .FileName
        If Trim(strFileName) <> "" Then
            mobjWB.SaveAs strFileName
            blnFinished = True
        Else
            strFileName = "False"
        End If
    End With
        
ErrHandle:
    mobjWB.Close
    Set mobjWS = Nothing
    Set mobjWB = Nothing
    If blnFinished Then
        MsgBox "��׼�ļ������Ѿ����ɣ�", vbInformation, gstrSysName
    ElseIf Trim(strFileName) <> "False" Then
        MsgBox "���ɱ�׼�ļ�����ʧ�ܣ�", vbInformation, gstrSysName
    End If
End Sub

Private Sub cmdProvider_Click()
    If Select��Ӧ��(Me, txtProvider, "") Then
        OS.PressKey vbKeyTab
    Else
        txtProvider.SetFocus
    End If
End Sub

Private Function CheckQualifications(ByVal intType As Integer, ByVal strInput As String) As Boolean
    'У�����ģ������̣���Ӧ����Ϣ������Ч��
    'intType��0�����ģ�1�������̣�2����Ӧ��
    'strInput���ַ���ʱΪ���ƣ�����ʱΪID
    Dim rsTmp As ADODB.Recordset
    Dim strMsgInfo As String
    Dim strMsgDate As String
    Dim dateCurrent As Date
    Dim strMsg As String
    
    Dim intCheckType As Integer
    Dim arrColumn
    Dim strCheck As String
    Dim strCheck_���� As String
    Dim strCheck_������ As String
    Dim strCheck_��Ӧ�� As String
    Dim n As Integer
    Dim strTmp As String
    
    On Error GoTo ErrHandle
    If strInput = "" Then
        CheckQualifications = True
        Exit Function
    End If
        
    '����У����Ŀ�ͷ�ʽ�ı����ʽ��У�鷽ʽ|���1,��Ŀ1,�Ƿ�У��;���1,��Ŀ2,�Ƿ�У��;���2,��Ŀ1,�Ƿ�У��;���2,��Ŀ2....
    strCheck = zlDatabase.GetPara("����У��", glngSys, mlngModule, "")
    
    '����Ĳ�����ʽ����ȷʱ�˳�
    If InStr(1, strCheck, "|") = 0 Then
        CheckQualifications = True
        Exit Function
    End If

    'ȡУ�鷽ʽ��0-����飻1�����ѣ�2����ֹ
    intCheckType = Val(Mid(strCheck, 1, InStr(1, strCheck, "|") - 1))
    
    '�����ʱ�˳�
    If intCheckType = 0 Then
        CheckQualifications = True
        Exit Function
    End If

    'ȡУ�����ݣ�
    strCheck = Mid(strCheck, InStr(1, strCheck, "|") + 1)

    If strCheck = "" Then
        CheckQualifications = True
        Exit Function
    End If

    '�ֱ�ȡ���ģ������̣���Ӧ����ҪУ�������
    strCheck = strCheck & ";"
    arrColumn = Split(strCheck, ";")
    For n = 0 To UBound(arrColumn)
        If arrColumn(n) <> "" Then
            If Split(arrColumn(n), ",")(0) = "����" And Split(arrColumn(n), ",")(2) = 1 Then
                strCheck_���� = IIf(strCheck_���� = "", "", strCheck_���� & ";") & Split(arrColumn(n), ",")(1)
            End If

            If Split(arrColumn(n), ",")(0) = "����������" And Split(arrColumn(n), ",")(2) = 1 Then
                strCheck_������ = IIf(strCheck_������ = "", "", strCheck_������ & ";") & Split(arrColumn(n), ",")(1)
            End If

            If Split(arrColumn(n), ",")(0) = "���Ĺ�Ӧ��" And Split(arrColumn(n), ",")(2) = 1 Then
                strCheck_��Ӧ�� = IIf(strCheck_��Ӧ�� = "", "", strCheck_��Ӧ�� & ";") & Split(arrColumn(n), ",")(1)
            End If
        End If
    Next
    
    '��У������ʱ�˳�
    If (intType = 0 And strCheck_���� = "") Or (intType = 1 And strCheck_������ = "") Or (intType = 2 And strCheck_��Ӧ�� = "") Then
        CheckQualifications = True
        Exit Function
    End If
    
    dateCurrent = CDate(Format(sys.Currentdate, "yyyy-mm-dd"))
    
    '����
    If intType = 0 Then
        gstrSQL = "Select ('[' || B.���� || ']' || B.����) AS ������Ϣ, A.���֤��, A.���֤��Ч�� " & _
            " From �շ���ĿĿ¼ B,�������� A " & _
            " Where B.ID = A.����ID And A.����ID = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "У����������", Val(strInput))
        
        If Not rsTmp.EOF Then
            If zlStr.Nvl(rsTmp!���֤��) = "" And InStr(strCheck_����, "���֤��") > 0 Then
                strTmp = rsTmp!������Ϣ & "��" & "�����֤��"
            End If
            
            If zlStr.Nvl(rsTmp!���֤��Ч��) <> "" Then
                If DateDiff("d", rsTmp!���֤��Ч��, dateCurrent) > 0 And InStr(strCheck_����, "���֤��Ч��") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!������Ϣ & "��", strTmp & ",") & "���֤�ѹ���"
                End If
            End If
        End If
    End If
    
    '������
    If intType = 1 Then
        gstrSQL = "Select ('[' || A.���� || ']' || A.����) AS ������, A.������ҵ���֤, A.������ҵ���֤Ч��,a.��Ӫ���֤, a.��Ӫ���֤Ч��, a.��ҵ����ִ��, a.��ҵ����ִ��Ч�� " & _
                        " From ���������� A " & _
                        " Where A.���� = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "У����������", strInput)
        
        If Not rsTmp.EOF Then
            If zlStr.Nvl(rsTmp!������ҵ���֤) = "" And InStr(strCheck_������ & ";", "������ҵ���֤" & ";") > 0 Then
                strTmp = rsTmp!������ & "��" & "��������ҵ���֤"
            End If
            
            If zlStr.Nvl(rsTmp!������ҵ���֤Ч��) <> "" Then
                If DateDiff("d", rsTmp!������ҵ���֤Ч��, dateCurrent) > 0 And InStr(strCheck_������, "������ҵ���֤Ч��") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!������ & "��", strTmp & ",") & "������ҵ���֤�ѹ���"
                End If
            End If
        End If
        
        If Not rsTmp.EOF Then
            If zlStr.Nvl(rsTmp!��Ӫ���֤) = "" And InStr(strCheck_������ & ";", "��Ӫ���֤" & ";") > 0 Then
                strTmp = rsTmp!������ & "��" & "�޾�Ӫ���֤"
            End If
            
            If zlStr.Nvl(rsTmp!��Ӫ���֤Ч��) <> "" Then
                If DateDiff("d", rsTmp!������ҵ���֤Ч��, dateCurrent) > 0 And InStr(strCheck_������ & ";", "��Ӫ���֤Ч��" & ";") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!������ & "��", strTmp & ",") & "��Ӫ���֤�ѹ���"
                End If
            End If
        End If
        If Not rsTmp.EOF Then
            If zlStr.Nvl(rsTmp!��ҵ����ִ��) = "" And InStr(strCheck_������ & ";", "��ҵ����ִ��" & ";") > 0 Then
                strTmp = rsTmp!������ & "��" & "����ҵ����ִ��"
            End If
            
            If zlStr.Nvl(rsTmp!��ҵ����ִ��Ч��) <> "" Then
                If DateDiff("d", rsTmp!������ҵ���֤Ч��, dateCurrent) > 0 And InStr(strCheck_������ & ";", "��ҵ����ִ��Ч��" & ";") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!������ & "��", strTmp & ",") & "��ҵ����ִ���ѹ���"
                End If
            End If
        End If
    End If
    
    '��Ӧ��
    If intType = 2 Then
        gstrSQL = "Select ('[' || ���� || ']' || ����) AS ��Ӧ��, ˰��ǼǺ�, ���֤��, ִ�պ�, ��Ȩ��, ������֤��, ������֤����, ҩ��ֱ�����, ҩ��ֱ�������, ���֤Ч��, ִ��Ч��, ��Ȩ�� " & _
            " From ��Ӧ�� " & _
            " Where (����ʱ�� Is Null Or ����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And ID = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ӧ����Ϣ", Val(strInput))
        
        strTmp = ""
        
        If Not rsTmp.EOF Then
            If zlStr.Nvl(rsTmp!˰��ǼǺ�) = "" And InStr(strCheck_��Ӧ��, "˰��ǼǺ�") > 0 Then
                strTmp = rsTmp!��Ӧ�� & "��" & "��˰��ǼǺ�"
            End If
            
            If zlStr.Nvl(rsTmp!���֤��) = "" And InStr(strCheck_��Ӧ��, "���֤��") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "�����֤��"
            End If
            
            If zlStr.Nvl(rsTmp!ִ�պ�) = "" And InStr(strCheck_��Ӧ��, "ִ�պ�") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "��ִ�պ�"
            End If
            
            If zlStr.Nvl(rsTmp!��Ȩ��) = "" And InStr(strCheck_��Ӧ��, "��Ȩ��") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "����Ȩ��"
            End If
            
            If zlStr.Nvl(rsTmp!������֤��) = "" And InStr(strCheck_��Ӧ��, "������֤��") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "��������֤��"
            End If
            
            If zlStr.Nvl(rsTmp!������֤����) <> "" Then
                If DateDiff("d", rsTmp!������֤����, dateCurrent) > 0 And InStr(strCheck_��Ӧ��, "������֤����") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "������֤���ѹ���"
                End If
            End If
            
            If zlStr.Nvl(rsTmp!ҩ��ֱ�����) = "" And InStr(strCheck_��Ӧ��, "ҩ��ֱ�����") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "��ҩ��ֱ�����"
            End If
            
            If zlStr.Nvl(rsTmp!ҩ��ֱ�������) <> "" Then
                If DateDiff("d", rsTmp!ҩ��ֱ�������, dateCurrent) > 0 And InStr(strCheck_��Ӧ��, "ҩ��ֱ�������") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "ҩ��ֱ������ѹ���"
                End If
            End If
            
            If zlStr.Nvl(rsTmp!���֤Ч��) <> "" Then
                If DateDiff("d", rsTmp!���֤Ч��, dateCurrent) > 0 And InStr(strCheck_��Ӧ��, "���֤Ч��") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "���֤�ѹ���"
                End If
            End If
            
            If zlStr.Nvl(rsTmp!ִ��Ч��) <> "" Then
                If DateDiff("d", rsTmp!ִ��Ч��, dateCurrent) > 0 And InStr(strCheck_��Ӧ��, "ִ��Ч��") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "ִ���ѹ���"
                End If
            End If
            
            If zlStr.Nvl(rsTmp!��Ȩ��) <> "" Then
                If DateDiff("d", rsTmp!ִ��Ч��, dateCurrent) > 0 And InStr(strCheck_��Ӧ��, "��Ȩ��") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "��Ȩ�ѹ���"
                End If
            End If
        End If
    End If
    
    '��ʾ���ֹ
    If strTmp <> "" Then
        If intCheckType = 1 Then
            If MsgBox("δͨ������У�飬�Ƿ������" & vbCrLf & strTmp, vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                CheckQualifications = True
                Exit Function
            Else
                Exit Function
            End If
        ElseIf intCheckType = 2 Then
            MsgBox "δͨ������У�飬������⣡" & vbCrLf & strTmp, vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    CheckQualifications = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function EntryPort(ByVal lngModule As Long, ByVal strStockInfo As String)
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    mlngModule = lngModule
    mlngStockID = Val(Split(strStockInfo, ";")(0))
    mstrStock = Split(strStockInfo, ";")(1)
    Caption = "�����ⲿ�ļ�(" & mstrStock & ")"
    
    Select Case mlngModule
    Case 1712
        Mstr_Cols = ";���ı���;��������;���;����;����;��������;Ч��;�������;���Ч��;����;��λ;�ɱ���;�ɱ����;��Ʊ��;��Ʊ����;��Ʊ���;��Ʒ��;"
        lblProvider.Caption = "��Ӧ��(&P)"
        cboIOType.Visible = False
    Case 1714
        Mstr_Cols = ";���ı���;��������;���;����;����;��������;Ч��;�������;���Ч��;����;��λ;�ɱ���;�ɱ����;��Ʒ��;"
        lblProvider.Caption = "�����(&I)"
        txtProvider.Visible = False
        cmdProvider.Visible = False
        cboIOType.Top = txtProvider.Top
        gstrSQL = "Select b.Id, b.���� From ҩƷ�������� A, ҩƷ������ B Where a.���id = b.Id And a.���� = 32"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        cboIOType.Clear
        Do While Not rsTmp.EOF
            cboIOType.AddItem rsTmp!����
            cboIOType.ItemData(cboIOType.NewIndex) = rsTmp!Id
            rsTmp.MoveNext
        Loop
        If rsTmp.RecordCount > 0 Then cboIOType.ListIndex = 0
    End Select
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SaveCard()
    Dim lngRow As Long
    Dim lngCol As Long
    
    On Error GoTo ErrHandle
    
    If mlngModule = 1712 Then
        If Val(txtProvider.Tag) = 0 Then
            MsgBox "δѡ��Ӧ�̣�", vbInformation, gstrSysName
            txtProvider.SetFocus
            Exit Sub
        End If
    ElseIf mlngModule = 1714 Then
        If cboIOType.ListIndex < 0 Then
            MsgBox "δѡ������࣡", vbInformation, gstrSysName
            cboIOType.SetFocus
            Exit Sub
        End If
    End If
    If vsfList.Rows = 1 Then Exit Sub
    vsfError.Rows = 1
    Call CheckData '����ʱ�������
    With vsfError
        If .Rows > 1 Then
            For lngRow = 1 To .Rows - 1
                If mbyt���뷽ʽ = 1 Then
                    If .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(2).Picture Or .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(1).Picture Then
                        MsgBox "��ȫ���뷽ʽ�£����ܴ����κβ��ϸ�����ݣ���������", vbInformation, gstrSysName
                        Exit Sub
                    End If
                Else
                    If .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(2).Picture Or .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(1).Picture Then
                        If MsgBox("�����ڲ��ϸ����ݣ��Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Exit Sub
                        Else
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
    End With
        
    '����
    Call ImportData
    Exit Sub
    
ErrHandle:
    If Not mobjWB Is Nothing Then
        mobjWB.Close
    End If
    Set mobjWB = Nothing
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    Dim rsTemp As ADODB.Recordset
    Dim cbrControl As CommandBarControl
    
    On Error GoTo ErrHandle
    
    Me.Height = 600 * 15
    Me.Width = 800 * 15
    
    Call InitComandbar
    Call InitControlPosition
    Call InitVSF
    ProCheck.Value = 100
    mintUnit = 1
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
        .FM_ɢװ���ۼ� = GetFmtString(0, g_�ۼ�)
    End With
    gstrSQL = "Select 1 From ��������˵�� Where ����id = [1] And �������� = '����ⷿ' And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ⷿ���ʲ�ѯ", mlngStockID)
    If rsTemp.RecordCount > 0 Then mblnVirtualStock = True
        
    If vsfList.Rows = 1 Then
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLOUTPUT, , True)
        cbrControl.Enabled = False
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLCHECK, , True)
        cbrControl.Enabled = False
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLSAVE, , True)
        cbrControl.Enabled = False
    End If
    
    Exit Sub

ErrHandle:
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub InitVSF()
    '��ʼ�����ؼ�
    With vsfList
        .Rows = 1
        .Cols = 17
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExNone   '�в�֧��������϶�
    End With
    
    With vsfError
        .Rows = 1
        .Cols = 4
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "����λ��"
        .ColWidth(1) = 2000
        .TextMatrix(0, 2) = "��������"
        .ColWidth(2) = 1500
        .TextMatrix(0, 3) = "����ԭ��"
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
'        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .ColWidth(0) = 300
        .ExtendLastCol = True '���һ�������
        .ExplorerBar = flexExNone   '�в�֧��������϶�
    End With
End Sub

Private Sub InitExcel()
    '��ʼ��Excel���
    Set mobjXLS = CreateObject("Excel.Application")
    mobjXLS.DisplayAlerts = False
End Sub

Private Function GetExcelData()
    '��ȡexcel������ݣ���������ʾ����������
    '����true-�д��� ����false-û�д���
    Dim strFileColumn As String '�ļ���������
    Dim lngRow As Long
    Dim lngCol As Long
    Dim intNum As Integer   '������
    Dim intCost As Integer  '�ɱ�����
    Dim dbl�ɱ���� As Double
    Dim bln�ɱ���� As Boolean
    Dim dbl��Ʊ��� As Double
    Dim bln��Ʊ��� As Double
    Dim blnNotNullRow As Boolean   '�����в��ǿ���
    Dim bln�������� As Boolean
    Dim blnЧ�� As Boolean
    Dim bln������� As Boolean
    Dim bln���Ч�� As Boolean
    Dim bln��Ʊ���� As Boolean
    
    Dim str�������� As String
    Dim strЧ�� As String
    Dim str������� As String
    Dim str���Ч�� As String
    Dim str��Ʊ���� As String
    
    On Error GoTo ErrHandle
    If txtFile.Text = "" Then Exit Function
    
    Set mobjWB = mobjXLS.Workbooks.Open(txtFile.Text)
    Set mobjWS = mobjWB.Sheets(1)
    If mobjWS Is Nothing Then Exit Function
    
    With mobjWS.UsedRange
        '��������˳����
        For lngCol = 1 To .Columns.count
            strFileColumn = strFileColumn & ";" & .cells(1, lngCol)
        Next
        For lngCol = 1 To UBound(Split(strFileColumn, ";"))
            If InStr(1, Mstr_Cols, ";" & Split(strFileColumn, ";")(lngCol) & ";") = 0 Then
                vsfError.Rows = vsfError.Rows + 1
                vsfError.TextMatrix(vsfError.Rows - 1, 1) = "1��"
                vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ͷ����"
                vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�����ļ���ͷ��" & Split(strFileColumn, ";")(lngCol) & "�������ڣ���ȷ��ͷӦ���ǡ�" & Mstr_Cols & "��������Ҫ�����Excel�ļ���"
                GetExcelData = True
            End If
        Next
        If mbyt���뷽ʽ = 1 Then
            If mbyt���ı��� = 1 Then
                If InStr(1, strFileColumn, ";���ı���;") = 0 Then
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = "1��"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ͷ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�����ı��롿�в����ڣ�������Ҫ�����Excel�ļ���"
                    GetExcelData = True
                End If
            End If
            If mbyt���� = 1 Then
                If InStr(1, strFileColumn, ";����;") = 0 Then
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = "1��"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ͷ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���������в����ڣ�������Ҫ�����Excel�ļ���"
                    GetExcelData = True
                End If
            End If
            If mbyt�ɱ��� = 1 Then
                If InStr(1, strFileColumn, ";�ɱ���;") = 0 Then
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = "1��"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ͷ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���ɱ��ۡ��в����ڣ�������Ҫ�����Excel�ļ���"
                    GetExcelData = True
                End If
            End If
            If mbyt�ɱ���� = 1 Then
                If InStr(1, strFileColumn, ";�ɱ����;") = 0 Then
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = "1��"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ͷ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���ɱ����в����ڣ�������Ҫ�����Excel�ļ���"
                    GetExcelData = True
                End If
            End If
        End If
        '������鲻��ͨ������Ҫ�޸ĵ����ļ�������˳����
        If GetExcelData = True Then
            Exit Function
        End If
        '��������������
        vsfList.Redraw = flexRDNone
        vsfList.Cols = .Columns.count + 1
        vsfList.Rows = 1
        For lngRow = 1 To .Rows.count
            vsfList.Rows = vsfList.Rows + 1
            For lngCol = 1 To .Columns.count
                vsfList.TextMatrix(lngRow - 1, lngCol) = .cells(lngRow, lngCol)
                If lngRow = 1 Then
                    vsfList.ColKey(lngCol) = .cells(lngRow, lngCol)
                End If
            Next
        Next
        '������ɾ��
        For lngRow = vsfList.Rows - 1 To 1 Step -1
            blnNotNullRow = True
            For lngCol = 1 To vsfList.Cols - 1
                If vsfList.TextMatrix(lngRow, lngCol) <> "" Then
                    blnNotNullRow = False
                End If
            Next
            '����ǿ��н���ɾ��
            If blnNotNullRow = True Then vsfList.RemoveItem lngRow
        Next
        Set mobjWS = Nothing
        Set mobjWB = Nothing
        mobjXLS.quit
        Set mobjXLS = Nothing
        With vsfList
            bln�ɱ���� = IIf(GetColumnPostation("�ɱ����") > 0, True, False)
            bln��Ʊ��� = IIf(GetColumnPostation("��Ʊ���") > 0, True, False)
            bln�������� = IIf(GetColumnPostation("��������") > 0, True, False)
            blnЧ�� = IIf(GetColumnPostation("Ч��") > 0, True, False)
            bln������� = IIf(GetColumnPostation("�������") > 0, True, False)
            bln���Ч�� = IIf(GetColumnPostation("���Ч��") > 0, True, False)
            bln��Ʊ���� = IIf(GetColumnPostation("��Ʊ����") > 0, True, False)
            For lngRow = 1 To .Rows - 1
                If bln�������� = True Then
                    str�������� = FormatDate(.TextMatrix(lngRow, .ColIndex("��������")))
                    .TextMatrix(lngRow, .ColIndex("��������")) = IIf(str�������� = "", .TextMatrix(lngRow, .ColIndex("��������")), str��������)
                End If
                If blnЧ�� = True Then
                    strЧ�� = FormatDate(.TextMatrix(lngRow, .ColIndex("Ч��")))
                    .TextMatrix(lngRow, .ColIndex("Ч��")) = IIf(strЧ�� = "", .TextMatrix(lngRow, .ColIndex("Ч��")), strЧ��)
                End If
                If bln������� = True Then
                    str������� = FormatDate(.TextMatrix(lngRow, .ColIndex("�������")))
                    .TextMatrix(lngRow, .ColIndex("�������")) = IIf(str������� = "", .TextMatrix(lngRow, .ColIndex("�������")), str�������)
                End If
                If bln���Ч�� = True Then
                    str���Ч�� = FormatDate(.TextMatrix(lngRow, .ColIndex("���Ч��")))
                    .TextMatrix(lngRow, .ColIndex("���Ч��")) = IIf(str���Ч�� = "", .TextMatrix(lngRow, .ColIndex("���Ч��")), str���Ч��)
                End If
                If bln��Ʊ���� = True Then
                    str��Ʊ���� = FormatDate(.TextMatrix(lngRow, .ColIndex("��Ʊ����")))
                    .TextMatrix(lngRow, .ColIndex("��Ʊ����")) = IIf(str��Ʊ���� = "", .TextMatrix(lngRow, .ColIndex("��Ʊ����")), FormatDate(str��Ʊ����))
                End If
                If bln�ɱ���� = True Then
                    dbl�ɱ���� = dbl�ɱ���� + Val(.TextMatrix(lngRow, .ColIndex("�ɱ����")))
                End If
                If bln��Ʊ��� = True Then
                    dbl��Ʊ��� = dbl��Ʊ��� + Val(.TextMatrix(lngRow, .ColIndex("��Ʊ���")))
                End If
            Next
            lblCollect.Caption = "�ɱ���" & Format(dbl�ɱ����, mFMT.FM_���) & "Ԫ          ��Ʊ��" & Format(dbl��Ʊ���, mFMT.FM_���) & "Ԫ"
        End With
        Call SetColumn
        Call CheckData
        
        vsfList.Redraw = flexRDDirect
    End With
    Exit Function
    
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function FormatDate(ByVal strDate As String) As String
    '���ܣ���ʽ�����ڣ������÷ֺ�(-)�ָ������ڸ�ʽ
    '���ظ�ʽ�����ֵ�����Ϊ����˵���������ڸ�ʽ����Ϊ��˵�������ڸ�ʽ
    Dim strYear, strMonth, strDay As String
    
    If LenB(StrConv(strDate, vbFromUnicode)) >= 8 Then
        If InStr(1, strDate, ".") > 0 Or InStr(1, strDate, "/") > 0 Or InStr(1, strDate, "-") > 0 Then
            strDate = Replace(strDate, ".", "")
            strDate = Replace(strDate, "/", "")
            strDate = Replace(strDate, "-", "")
        End If
        strYear = Mid(strDate, 1, 4)
        If LenB(StrConv(strDate, vbFromUnicode)) < 8 Then
            strMonth = Mid(strDate, 5, 1)
        Else
            strMonth = Mid(strDate, 5, 2)
        End If
        If LenB(StrConv(strDate, vbFromUnicode)) < 8 Then
            strDay = Mid(strDate, 6, 1)
        Else
            strDay = Mid(strDate, 7, 2)
        End If
        If IsNumeric(strYear) = True And IsNumeric(strMonth) = True And IsNumeric(strDay) = True Then
            FormatDate = strYear & "-" & strMonth & "-" & strDay
        End If
    Else
        FormatDate = ""
    End If
End Function

Private Function GetColumnPostation(ByVal strColumn As String) As Integer
    '��ȡ��λ�ú��ж��Ƿ����
    '���� strcolumn-���������
    '����ֵ :���ش�����λ�� 0-û���ҵ� >0�ҵ���
    Dim lngRow As Long
    Dim lngCol As Long
    
    With vsfList
        For lngCol = 1 To .Cols - 1
            If .TextMatrix(0, lngCol) = strColumn Then
                GetColumnPostation = lngCol
                Exit Function
            End If
        Next
        GetColumnPostation = 0
    End With
End Function

Private Sub CheckData()
    '������ݺϷ���
    Dim lngRow As Long
    Dim lngCol As Long
    Dim rsTemp As ADODB.Recordset
    Dim dbl�ɱ��� As Double
    Dim dbl���� As Double
    Dim dbl�ɱ���� As Double
    Dim dbl��Ʊ��� As Double
    Dim cbrControl As CommandBarControl
    Dim strTemp As String
    
    On Error GoTo ErrHandle
    
    Call ParseParameter
    vsfError.Rows = 1
    With vsfList
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = vbBlack '�����óɺ�ɫ
        ProCheck.Value = 0
        For lngRow = 1 To .Rows - 1
            .TextMatrix(lngRow, 0) = lngRow '����б�
            '���ı���
            If GetColumnPostation("���ı���") > 0 Then
                If .TextMatrix(lngRow, .ColIndex("���ı���")) = "" Then
                    .Cell(flexcpForeColor, lngRow, .ColIndex("���ı���"), lngRow, .ColIndex("���ı���")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt���ı��� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С����ı��롿��"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�����ı��롿��Ϊ�գ���������"
                Else
                    gstrSQL = "Select 1 From �շ���ĿĿ¼ Where ��� = '4' And ���� =[1] And Rownum < 2"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "������֤", .TextMatrix(lngRow, .ColIndex("���ı���")))
                    If rsTemp.RecordCount = 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("���ı���"), lngRow, .ColIndex("���ı���")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt���ı��� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С����ı��롿��"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�����ı��롿�б��벻���ڣ���������"
                    End If
                End If
            End If
            '��������
            If GetColumnPostation("��������") > 0 Then
                .TextMatrix(lngRow, .ColIndex("��������")) = Trim(.TextMatrix(lngRow, .ColIndex("��������")))
            End If
            '���
            If GetColumnPostation("���") > 0 Then
                .TextMatrix(lngRow, .ColIndex("���")) = Trim(.TextMatrix(lngRow, .ColIndex("���")))
            End If
            '����
            If .TextMatrix(lngRow, .ColIndex("����")) = "" Then
                .Cell(flexcpForeColor, lngRow, .ColIndex("����"), lngRow, .ColIndex("����")) = vbRed
                vsfError.Rows = vsfError.Rows + 1
                vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt���� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С���������"
                vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ����"
                vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�����������Ǳ������Ҵ����㣬��������"
            Else
                If Not IsNumeric(.TextMatrix(lngRow, .ColIndex("����"))) Then
                    .Cell(flexcpForeColor, lngRow, .ColIndex("����"), lngRow, .ColIndex("����")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt���� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С���������"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ʽ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "����������ӦΪ�������ͣ���������"
                Else
                    .TextMatrix(lngRow, .ColIndex("����")) = Format(.TextMatrix(lngRow, .ColIndex("����")), mFMT.FM_����)
                    If Val(.TextMatrix(lngRow, .ColIndex("����"))) > 9999999999# Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("����"), lngRow, .ColIndex("����")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt���� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С���������"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "���ݴ���"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���������д���9999999999#����������"
                    End If
                End If
            End If
            '�ɱ���
            If .TextMatrix(lngRow, .ColIndex("�ɱ���")) = "" Then
                .Cell(flexcpForeColor, lngRow, .ColIndex("�ɱ���"), lngRow, .ColIndex("�ɱ���")) = vbRed
                vsfError.Rows = vsfError.Rows + 1
                vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt�ɱ��� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С��ɱ��ۡ���"
                vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ��ʾ"
                vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���ɱ��ۡ���Ϊ���ˣ�"
            Else
                If Not IsNumeric(.TextMatrix(lngRow, .ColIndex("�ɱ���"))) Then
                    .Cell(flexcpForeColor, lngRow, .ColIndex("�ɱ���"), lngRow, .ColIndex("�ɱ���")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt�ɱ��� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С��ɱ��ۡ���"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ʽ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���ɱ��ۡ���ӦΪ�������ͣ���������"
                Else
                    .TextMatrix(lngRow, .ColIndex("�ɱ���")) = Format(.TextMatrix(lngRow, .ColIndex("�ɱ���")), mFMT.FM_�ɱ���)
                    If Val(.TextMatrix(lngRow, .ColIndex("�ɱ���"))) > 999999999 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("�ɱ���"), lngRow, .ColIndex("�ɱ���")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt�ɱ��� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С��ɱ��ۡ���"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "���ݴ���"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���ɱ��ۡ��д���999999999����������"
                    End If
                End If
            End If
            '�ɱ����
            If .TextMatrix(lngRow, .ColIndex("�ɱ����")) = "" Then
                .Cell(flexcpForeColor, lngRow, .ColIndex("�ɱ����"), lngRow, .ColIndex("�ɱ����")) = vbRed
                vsfError.Rows = vsfError.Rows + 1
                vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt�ɱ���� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С��ɱ�����"
                vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ��ʾ"
                vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���ɱ���Ϊ���ˣ�"
            Else
                If Not IsNumeric(.TextMatrix(lngRow, .ColIndex("�ɱ����"))) Then
                    .Cell(flexcpForeColor, lngRow, .ColIndex("�ɱ����"), lngRow, .ColIndex("�ɱ����")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt�ɱ���� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С��ɱ�����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ʽ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���ɱ�����ӦΪ�������ͣ���������"
                Else
                    .TextMatrix(lngRow, .ColIndex("�ɱ����")) = Format(.TextMatrix(lngRow, .ColIndex("�ɱ����")), mFMT.FM_���)
                    If Val(.TextMatrix(lngRow, .ColIndex("�ɱ����"))) > 9999999999# Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("�ɱ����"), lngRow, .ColIndex("�ɱ����")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt�ɱ���� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С��ɱ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ʽ����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���ɱ����д���9999999999����������"
                    End If
                End If
            End If
            '��Ʊ���
            If GetColumnPostation("��Ʊ���") > 0 Then
                If .TextMatrix(lngRow, .ColIndex("��Ʊ���")) = "" Then
                    .Cell(flexcpForeColor, lngRow, .ColIndex("��Ʊ���"), lngRow, .ColIndex("��Ʊ���")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt��Ʊ��� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С���Ʊ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ��ʾ"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "����Ʊ��Ϊ���ˣ�"
                Else
                    If Not IsNumeric(.TextMatrix(lngRow, .ColIndex("��Ʊ���"))) Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("��Ʊ���"), lngRow, .ColIndex("��Ʊ���")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt��Ʊ��� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С���Ʊ����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ʽ����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "����Ʊ����ӦΪ�������ͣ���������"
                    Else
                        .TextMatrix(lngRow, .ColIndex("��Ʊ���")) = Format(.TextMatrix(lngRow, .ColIndex("��Ʊ���")), mFMT.FM_���)
                        If Val(.TextMatrix(lngRow, .ColIndex("��Ʊ���"))) > 999999999999# Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("��Ʊ���"), lngRow, .ColIndex("��Ʊ���")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt��Ʊ��� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С���Ʊ����"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "���ݴ���"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "����Ʊ���д���999999999999����������"
                        End If
                    End If
                End If
            End If
            '����*�ɱ���=�ɱ����
            dbl�ɱ��� = Format(Val(.TextMatrix(lngRow, .ColIndex("�ɱ���"))), mFMT.FM_�ɱ���)
            dbl���� = Format(Val(.TextMatrix(lngRow, .ColIndex("����"))), mFMT.FM_����)
            dbl�ɱ���� = Format(Val(.TextMatrix(lngRow, .ColIndex("�ɱ����"))), mFMT.FM_���)
            If dbl�ɱ��� * dbl���� <> dbl�ɱ���� Then
                .Cell(flexcpForeColor, lngRow, .ColIndex("�ɱ����"), lngRow, .ColIndex("�ɱ����")) = vbRed
                vsfError.Rows = vsfError.Rows + 1
                vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytNumCost = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С��ɱ�����"
                vsfError.TextMatrix(vsfError.Rows - 1, 2) = "���ݴ���"
                vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�ɱ���*����<>�ɱ�����������"
            End If
            '��Ʊ���=�ɱ����
            If GetColumnPostation("��Ʊ���") > 0 Then
                dbl��Ʊ��� = Format(Val(.TextMatrix(lngRow, .ColIndex("��Ʊ���"))), mFMT.FM_���)
                If dbl��Ʊ��� <> dbl�ɱ���� Then
                    .Cell(flexcpForeColor, lngRow, .ColIndex("��Ʊ���"), lngRow, .ColIndex("��Ʊ���")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytInvoiceCost = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С���Ʊ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "������ʾ"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��Ʊ���<>�ɱ������飡"
                End If
            End If
            '�ļ��ɱ���=HIS�ɱ���
            gstrSQL = "Select Nvl(c.�ɱ���, 0) �ɱ���, Nvl(c.����ϵ��, 1) As ����ϵ��" & vbNewLine & _
                    "From �շ���ĿĿ¼ B, �������� C" & vbNewLine & _
                    "Where b.Id = c.����id And b.���� = [1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ɱ���", .TextMatrix(lngRow, .ColIndex("���ı���")))
            If rsTemp.RecordCount <> 0 Then
                If Format(Val(.TextMatrix(lngRow, .ColIndex("�ɱ���"))), mFMT.FM_�ɱ���) <> Format(rsTemp!�ɱ��� * rsTemp!����ϵ��, mFMT.FM_�ɱ���) Then
                    .Cell(flexcpForeColor, lngRow, .ColIndex("�ɱ���"), lngRow, .ColIndex("�ɱ���")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytExcelCost = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С��ɱ��ۡ���"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "������ʾ"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�ĵ����ɱ��ۡ�" & Format(Val(.TextMatrix(lngRow, .ColIndex("�ɱ���"))), mFMT.FM_�ɱ���) & "��HISϵͳ����С��ɱ��ۡ�" & Format(rsTemp!�ɱ��� * rsTemp!����ϵ��, mFMT.FM_�ɱ���) & "���ȣ�"
                End If
            End If
            'Ч��
            If GetColumnPostation("Ч��") > 0 Then
                If .TextMatrix(lngRow, .ColIndex("Ч��")) <> "" Then
                    If Not IsDate(.TextMatrix(lngRow, .ColIndex("Ч��"))) Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("Ч��"), lngRow, .ColIndex("Ч��")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytЧ�� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�Ч�ڡ���"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ʽ����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��Ч�ڡ���ӦΪ���ڸ�ʽ����3000-01-01��3000/01/01����30000101��"
                    End If
                End If
            End If
             
            '�������
            If GetColumnPostation("�������") > 0 Then
                If .TextMatrix(lngRow, .ColIndex("�������")) <> "" Then
                    If Not IsDate(.TextMatrix(lngRow, .ColIndex("�������"))) Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("�������"), lngRow, .ColIndex("�������")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt������� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�������ڡ���"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ʽ����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��������ڡ���ӦΪ���ڸ�ʽ����3000-01-01��3000/01/01����30000101��"
                    End If
                End If
            End If
            '���Ч��
            If GetColumnPostation("���Ч��") > 0 Then
                If .TextMatrix(lngRow, .ColIndex("���Ч��")) <> "" Then
                    If Not IsDate(.TextMatrix(lngRow, .ColIndex("���Ч��"))) Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("���Ч��"), lngRow, .ColIndex("���Ч��")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt���Ч�� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С����Ч�ڡ���"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ʽ����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�����Ч�ڡ���ӦΪ���ڸ�ʽ����3000-01-01��3000/01/01����30000101��"
                    End If
                End If
            End If
            '��Ʊ����
            If GetColumnPostation("��Ʊ����") > 0 Then
                If .TextMatrix(lngRow, .ColIndex("��Ʊ����")) <> "" Then
                    If Not IsDate(.TextMatrix(lngRow, .ColIndex("��Ʊ����"))) Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("��Ʊ����"), lngRow, .ColIndex("��Ʊ����")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt��Ʊ���� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С���Ʊ���ڡ���"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ʽ����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "����Ʊ���ڡ���ӦΪ�����ڸ�ʽ����3000-01-01��3000/01/01����30000101��"
                    End If
                End If
            End If
            If GetColumnPostation("��������") > 0 Then
                If .TextMatrix(lngRow, .ColIndex("��������")) <> "" Then
                    If Not IsDate(.TextMatrix(lngRow, .ColIndex("��������"))) Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("��������"), lngRow, .ColIndex("��������")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt�������� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С��������ڡ���"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ʽ����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���������ڡ���ӦΪ���ڸ�ʽ����3000-01-01��3000/01/01����30000101��"
                    End If
                End If
            End If
            '���洢�ⷿ
            If .TextMatrix(lngRow, .ColIndex("���ı���")) <> "" Then
                gstrSQL = "Select 1" & vbNewLine & _
                            "From �շ���ĿĿ¼ A, �շ�ִ�п��� B" & vbNewLine & _
                            "Where a.Id = b.�շ�ϸĿid And b.ִ�п���id = [1] And a.��� = '4' And a.���� = [2] and rownum<2"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�洢�ⷿ", mlngStockID, .TextMatrix(lngRow, .ColIndex("���ı���")))
                If rsTemp.RecordCount = 0 Then
                    .Cell(flexcpForeColor, lngRow, .ColIndex("���ı���"), lngRow, .ColIndex("���ı���")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt�洢�ⷿ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С����ı��롿��"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "���ݴ���"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "������δ��" & mstrStock & "�ⷿ�����ô洢״̬����������Ŀ¼�е����洢״̬��"
                End If
                '��ֵ������Ҫ�����������
                If mblnVirtualStock = True Then
                    gstrSQL = "Select b.��ֵ����, b. ���ٲ���, b.��������, b.���÷���" & vbNewLine & _
                        "From �շ���ĿĿ¼ A, �������� B" & vbNewLine & _
                        "Where a.Id = b.����id And a.��� = '4' And a.���� = [1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�洢�ⷿ", .TextMatrix(lngRow, .ColIndex("���ı���")))
                        If (zlStr.Nvl(rsTemp!��ֵ����, 0) = 0 Or zlStr.Nvl(rsTemp!���ٲ���, 0) = 0 Or zlStr.Nvl(rsTemp!��������, 0) = 0 Or zlStr.Nvl(rsTemp!���÷���, 0) = 0) Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("���ı���"), lngRow, .ColIndex("���ı���")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt����ⷿ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С����ı��롿��"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "���ݴ���"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "����ⷿ���ı����Ǹ�ֵ���ϡ��������á����ٲ��ˡ����÷������뵽����Ŀ¼���޸ĸ��������ԣ�"
                        End If
                End If
            End If
            '��Ʒ������
            If GetColumnPostation("��Ʒ��") > 0 Then
                If .TextMatrix(lngRow, .ColIndex("��Ʒ��")) <> "" Then
                    gstrSQL = "Select 1" & vbNewLine & _
                                "From ҩƷ�շ���¼ A, �շ���ĿĿ¼ B" & vbNewLine & _
                                "Where a.ҩƷid = b.Id And b.���� = [1] And b.��� = '4' And a.�ⷿid = [2] And a.��Ʒ���� =[3] And Rownum < 2"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ʒ����", .TextMatrix(lngRow, .ColIndex("���ı���")), mlngStockID, .TextMatrix(lngRow, .ColIndex("��Ʒ��")))
                    If rsTemp.RecordCount <> 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("��Ʒ��"), lngRow, .ColIndex("��Ʒ��")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt��Ʒ���� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С���Ʒ�롿��"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "���ݴ���"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��Ʒ�����ظ������޸ģ�"
                    End If
                End If
            End If
            If ProCheck.Value + 100 / (vsfList.Rows - 1) >= 100 Then
                ProCheck.Value = 100
            Else
                ProCheck.Value = ProCheck.Value + 100 / (vsfList.Rows - 1)
            End If
        Next
    End With
    Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLSAVE, , True)
    cbrControl.Enabled = True
    With vsfError
        If .Rows > 1 Then
            If mbyt���뷽ʽ = 0 Then
                cbrControl.Enabled = True
            Else
                For lngRow = 1 To .Rows - 1
                    If .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(1).Picture Or .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(2).Picture Then
                        cbrControl.Enabled = False
                        Exit For
                    End If
                Next
            End If
        End If
    End With
    If vsfList.Rows > 1 Then
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLOUTPUT, , True)
        cbrControl.Enabled = True
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLCHECK, , True)
        cbrControl.Enabled = True
    End If
    Exit Sub
    
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetColumn()
    '�п�ȶ��뷽ʽ����
    Dim lngRow As Long
    Dim lngCol As Long
    
    With vsfList
        For lngCol = 1 To .Cols - 1
            Select Case .TextMatrix(0, lngCol)
                Case "���ı���", "��������", "���", "����", "����", "��Ʒ��", "��������", "Ч��", "���Ч��", "�������", "��Ʊ����"
                    .ColAlignment(lngCol) = flexAlignLeftCenter
                Case "����", "�ɱ���", "�ɱ����", "��Ʊ���"
                    .ColAlignment(lngCol) = flexAlignRightCenter
                Case Else
                    .ColAlignment(lngCol) = flexAlignRightCenter
            End Select
        Next
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColWidth(0) = 300
        If GetColumnPostation("���ı���") > 0 Then
            .ColWidth(.ColIndex("���ı���")) = 2000
        End If
        If GetColumnPostation("��������") > 0 Then
            .ColWidth(.ColIndex("��������")) = 1000
        End If
        If GetColumnPostation("Ч��") > 0 Then
            .ColWidth(.ColIndex("Ч��")) = 1000
        End If
        If GetColumnPostation("���Ч��") > 0 Then
            .ColWidth(.ColIndex("���Ч��")) = 1000
        End If
        If GetColumnPostation("�������") > 0 Then
            .ColWidth(.ColIndex("�������")) = 1000
        End If
        If GetColumnPostation("��Ʊ����") > 0 Then
            .ColWidth(.ColIndex("��Ʊ����")) = 1000
        End If
    End With
End Sub

Private Sub InitControlPosition()
    '�ؼ�λ�ÿ���
    On Error Resume Next
    
    lblFile.Move 100, 600
    txtFile.Move lblFile.Left + lblFile.Width + 20, lblFile.Top - 40, Me.ScaleWidth - (cmdFile.Width + txtFile.Left)
    cmdFile.Move txtFile.Left + txtFile.Width, txtFile.Top - 30
    lblProvider.Move lblFile.Left, lblFile.Top + lblFile.Height + 200
    txtProvider.Move txtFile.Left, lblProvider.Top - 50, txtFile.Width
    
    If mlngModule = 1712 Then
        cboIOType.Move txtFile.Left, lblProvider.Top, txtFile.Width
        cmdProvider.Move txtProvider.Left + txtProvider.Width - 10, cboIOType.Top - 60
    Else
        cboIOType.Move txtFile.Left, lblProvider.Top, txtFile.Width + cmdFile.Width
        cmdProvider.Visible = False
    End If
    vsfList.Move lblFile.Left, txtProvider.Top + txtProvider.Height + 50, Me.Width - lblFile.Left - vsfList.Left - 120, ((Me.ScaleHeight - vsfList.Top) / 4) * 3
    picSplit.Move lblFile.Left, vsfList.Top + vsfList.Height + 50, Me.ScaleWidth - lblFile.Left - vsfList.Left
    lblCollect.Width = picSplit.Width
    
    vsfError.Move lblFile.Left, picSplit.Top + picSplit.Height + 50, Me.ScaleWidth - lblFile.Left, Me.ScaleHeight - picSplit.Top - picSplit.Height - 100
    ProCheck.Move lblFile.Left, (vsfError.Top + vsfError.Height) / 2, Me.ScaleWidth - lblFile.Left - ProCheck.Left
    ProCheck.Visible = False
End Sub

Private Sub Form_Resize()
    Call InitControlPosition
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjWS = Nothing
    Set mobjWB = Nothing
    Set mobjXLS = Nothing
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    With picSplit
        If .Top + y < 2000 Then Exit Sub
        If .Top + y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + y
    End With

    With vsfList
        .Height = picSplit.Top - .Top
    End With
    
    With vsfError
        .Top = picSplit.Top + picSplit.Height + 100
        .Height = ScaleHeight - .Top
    End With
    Me.Refresh
End Sub

Private Sub txtFile_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(txtFile.Text) <> "" Then OS.PressKey vbKeyTab
End Sub

Private Sub txtFile_LostFocus()
'    If Trim(txtFile.Text) <> "" Then OS.PressKey vbKeyTab
End Sub

Private Sub txtProvider_Change()
    txtProvider.Tag = ""
End Sub

Private Sub txtProvider_GotFocus()
    OS.OpenIme False
    zlControl.TxtSelAll txtProvider
End Sub

Private Sub txtProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Val(txtProvider.Tag) <> 0 Then OS.PressKey vbKeyTab: Exit Sub
    If Val(txtProvider.Tag) = 0 And Trim(txtProvider.Text) = "" Then Exit Sub
    If Select��Ӧ��(Me, txtProvider, Trim(txtProvider.Text)) = False Then Exit Sub
    OS.PressKey vbKeyTab
End Sub

Private Sub txtProvider_LostFocus()
    If txtProvider.Text = "" Then
        txtProvider.Tag = "0"
        Exit Sub
    End If
End Sub

Private Sub ImportData()
'���ⲿ�ļ������ݵ��뵽ҩƷ�շ���¼����
    Dim lngCol As Long, lngRow As Long
    Dim lngMaterialID As Long
    Dim dblTransVal As Double, dblAddRate As Double
    Dim dblSalePrice As Double, dblSale As Double, dblCostPrice As Double, dblCost As Double, dblCurSale As Double
    Dim dblQTY As Double
    Dim arrCols As Variant
    Dim strTmp As String, strInsert As String, strNo As String, strMess As String
    Dim rsTmp As ADODB.Recordset
    Dim strPlaceProduction As String, strPackageUnit As String
    Dim bytLotPrice As Byte
    Dim blnLot As Boolean, blnOnce As Boolean
    Dim blnTran As Boolean  '��¼�����Ƿ�ʼ��
    
    On Error GoTo ErrHandle
    
    With vsfList
        '���ݺ�
        Select Case mlngModule
            Case 1712
                strNo = sys.GetNextNo(68, mlngStockID)
            Case 1714
                strNo = sys.GetNextNo(70, mlngStockID)
        End Select
        '��֯����
        
        On Error GoTo ErrHandle
        blnTran = True
        gcnOracle.BeginTrans
        For lngRow = 1 To .Rows - 1
            '������ı���
            gstrSQL = "select a.ID, a.�Ƿ���, 1/(1-b.ָ�������/100)-1 �ӳ���, b.����ϵ��, b.�ⷿ����" & _
                      ", b.һ���Բ���, b.��ֵ����, b. ���ٲ���, b.��������, b.���÷���, b.��װ��λ, c.�ּ� " & _
                      "from �շ���ĿĿ¼ a, �������� b, �շѼ�Ŀ c " & _
                      "where a.ID=b.����ID and a.ID=c.�շ�ϸĿID and a.����=[1] and a.���='4' " & _
                      " and a.����ʱ��>=to_date('3000-1-1','yyyy-mm-dd') and c.��ֹ����=to_date('3000-1-1','yyyy-mm-dd') " & _
                      GetPriceClassString("C")
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "" & .TextMatrix(lngRow, .ColIndex("���ı���")))
            '����ⷿ�Ŀⷿ�������[��ֵ����][���ٲ���][��������][���÷���]����
            If rsTmp.RecordCount > 0 Then
                lngMaterialID = rsTmp!Id
                dblCurSale = IIf(IsNull(rsTmp!�ּ�), 0, rsTmp!�ּ�)
                dblTransVal = IIf(IsNull(rsTmp!����ϵ��), 1, rsTmp!����ϵ��)
                dblAddRate = IIf(IsNull(rsTmp!�ӳ���), 0, rsTmp!�ӳ���)
                bytLotPrice = IIf(IsNull(rsTmp!�Ƿ���), 0, rsTmp!�Ƿ���)
                blnLot = IIf(IsNull(rsTmp!�ⷿ����), 0, rsTmp!�ⷿ����)
                blnOnce = IIf(IsNull(rsTmp!һ���Բ���), 0, rsTmp!һ���Բ���)
                strPackageUnit = IIf(IsNull(rsTmp!��װ��λ), "", rsTmp!��װ��λ)
                
                If mlngModule = 1712 Then
                    strInsert = "zl_�����⹺_INSERT("
                    'NO
                    strInsert = strInsert & "'" & strNo & "',"
                    '���
                    strInsert = strInsert & lngRow - 1 & ","
                    '�ⷿID
                    strInsert = strInsert & mlngStockID & ","
                    '��Ӧ��ID
                    strInsert = strInsert & txtProvider.Tag & ","
                    '����ID
                    strInsert = strInsert & lngMaterialID & ","
                    '����
                    If GetColumnPostation("����") > 0 Then
                        strTmp = UCase(Trim(.TextMatrix(lngRow, .ColIndex("����"))))
                    Else
                        strTmp = ""
                    End If
                    strInsert = strInsert & IIf(strTmp = "", "null", "'" & strTmp & "'") & ","
                    '����
                    If GetColumnPostation("����") > 0 Then
                        strTmp = Trim(.TextMatrix(lngRow, .ColIndex("����")))
                    Else
                        strTmp = ""
                    End If
                    strInsert = strInsert & IIf(strTmp = "", "null", "'" & strTmp & "'") & ","
                    '��������
                    If GetColumnPostation("��������") > 0 Then
                        strTmp = IIf(IsDate(Trim(.TextMatrix(lngRow, .ColIndex("��������")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("��������"))), "")
                    Else
                        strTmp = ""
                    End If
                    If IsNumeric(strTmp) Then
                        strTmp = TranNumToDate(strTmp)
                    End If
                    strInsert = strInsert & IIf(strTmp = "", "null", "to_date('" & Format(strTmp, "yyyy-mm-dd") & "','YYYY-MM-DD')") & ","
                    'Ч��
                    If blnLot Then  '�ⷿ��������Ч��
                        If GetColumnPostation("Ч��") > 0 Then
                            strTmp = IIf(IsDate(Trim(.TextMatrix(lngRow, .ColIndex("Ч��")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("Ч��"))), "")
                        Else
                            strTmp = ""
                        End If
                        If IsNumeric(strTmp) Then
                            strTmp = TranNumToDate(strTmp)
                        End If
                        strInsert = strInsert & IIf(strTmp = "", "null", "to_date('" & Format(strTmp, "yyyy-mm-dd") & "','YYYY-MM-DD')") & ","
                    Else
                        strInsert = strInsert & "null,"
                    End If
                    If blnOnce Then
                        '�������
                        If GetColumnPostation("�������") > 0 Then
                            strTmp = IIf(IsDate(Trim(.TextMatrix(lngRow, .ColIndex("�������")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("�������"))), "")
                        Else
                            strTmp = ""
                        End If
                        If IsNumeric(strTmp) Then
                            strTmp = TranNumToDate(strTmp)
                        End If
                        strInsert = strInsert & IIf(strTmp = "", "null", "to_date('" & Format(strTmp, "yyyy-mm-dd") & "','YYYY-MM-DD')") & ","
                        '���Ч��
                        If GetColumnPostation("���Ч��") > 0 Then
                            strTmp = IIf(IsDate(Trim(.TextMatrix(lngRow, .ColIndex("���Ч��")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("���Ч��"))), "")
                        Else
                            strTmp = ""
                        End If
                        If IsNumeric(strTmp) Then
                            strTmp = TranNumToDate(strTmp)
                        End If
                        strInsert = strInsert & IIf(strTmp = "", "null", "to_date('" & Format(strTmp, "yyyy-mm-dd") & "','YYYY-MM-DD')") & ","
                    Else
                        strInsert = strInsert & "null,null,"
                    End If
                    '����
                    If GetColumnPostation("����") > 0 Then
                        strTmp = IIf(IsNumeric(Trim(.TextMatrix(lngRow, .ColIndex("����")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("����"))), "")
                    Else
                        strTmp = ""
                    End If
                    dblQTY = IIf(strTmp = "", 0, strTmp)
                    strInsert = strInsert & GetFormat(dblQTY * dblTransVal, g_С��λ��.obj_���С��.����С��) & ","
                    '�ɱ���
                    If GetColumnPostation("�ɱ���") > 0 Then
                        strTmp = IIf(IsNumeric(Trim(.TextMatrix(lngRow, .ColIndex("�ɱ���")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("�ɱ���"))), "")
                    Else
                        strTmp = ""
                    End If
                    dblCostPrice = CDbl(IIf(strTmp = "", 0, strTmp))
                    strInsert = strInsert & GetFormat(dblCostPrice / IIf(dblTransVal = 0, 1, dblTransVal), g_С��λ��.obj_���С��.�ɱ���С��) & ","
                    '�ɱ����
                    If GetColumnPostation("�ɱ����") > 0 Then
                        strTmp = IIf(IsNumeric(Trim(.TextMatrix(lngRow, .ColIndex("�ɱ����")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("�ɱ����"))), "")
                    Else
                        strTmp = ""
                    End If
                    dblCost = CDbl(IIf(strTmp = "", 0, strTmp))
                    strInsert = strInsert & GetFormat(dblCost, g_С��λ��.obj_���С��.���С��) & ","
                    '����
                    strInsert = strInsert & "100,"
                    '�ۼ�
                    dblSalePrice = IIf(bytLotPrice = 1, dblCostPrice / IIf(dblTransVal = 0, 1, dblTransVal) * (dblAddRate + 1), dblCurSale)
                    strInsert = strInsert & GetFormat(dblSalePrice, g_С��λ��.obj_���С��.���ۼ�С��) & ","
                    '�ۼ۽��
                    dblSale = GetFormat(dblQTY * dblTransVal, g_С��λ��.obj_���С��.����С��) * GetFormat(dblSalePrice, g_С��λ��.obj_���С��.���ۼ�С��)
                    strInsert = strInsert & GetFormat(dblSale, g_С��λ��.obj_���С��.���С��) & ","
                    '���
                    strInsert = strInsert & GetFormat(dblSale, g_С��λ��.obj_���С��.���С��) - GetFormat(dblCost, g_С��λ��.obj_���С��.���С��) & ","
                    '���۲�ۣ�ժҪ��ע��֤��
                    strInsert = strInsert & "null,null,null,"
                    '������
                    strInsert = strInsert & "'" & gstrUserName & "',"
                    '�������
                    strInsert = strInsert & "null,"
                    '��Ʊ��
                    If GetColumnPostation("��Ʊ��") > 0 Then
                        strTmp = Trim(.TextMatrix(lngRow, .ColIndex("��Ʊ��")))
                    Else
                        strTmp = ""
                    End If
                    strInsert = strInsert & IIf(strTmp = "", "null", "'" & strTmp & "'") & ","
                    '��Ʊ����
                    If GetColumnPostation("��Ʊ����") > 0 Then
                        strTmp = IIf(IsDate(Trim(.TextMatrix(lngRow, .ColIndex("��Ʊ����")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("��Ʊ����"))), "")
                    Else
                        strTmp = ""
                    End If
                    If IsNumeric(strTmp) Then
                        strTmp = TranNumToDate(strTmp)
                    End If
                    strInsert = strInsert & IIf(strTmp = "", "null", "to_date('" & Format(strTmp, "yyyy-mm-dd") & "','YYYY-MM-DD')") & ","
                    '��Ʊ���
                    If GetColumnPostation("��Ʊ���") > 0 Then
                        strTmp = IIf(IsNumeric(Trim(.TextMatrix(lngRow, .ColIndex("��Ʊ���")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("��Ʊ���"))), "")
                    Else
                        strTmp = ""
                    End If
                    strInsert = strInsert & IIf(strTmp = "", 0, strTmp) & ","
                    '��������
                    strInsert = strInsert & "to_date('" & Format(Now(), "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd HH24:MI:SS'),"
                    '�˲��ˣ��˲����ڣ����Σ��˻�����ֵ����
                    strInsert = strInsert & "null,null,null,1,null,"
                    '��Ʒ��
                    If GetColumnPostation("��Ʒ��") > 0 Then
                        strTmp = Trim(.TextMatrix(lngRow, .ColIndex("��Ʒ��")))
                    Else
                        strTmp = ""
                    End If
                    strInsert = strInsert & IIf(strTmp = "", "null)", "'" & strTmp & "')")
                ElseIf mlngModule = 1714 Then
                    strInsert = "zl_�����������_INSERT("
                    'No
                    strInsert = strInsert & "'" & strNo & "',"
                    '���
                    strInsert = strInsert & lngRow - 1 & ","
                    '�ⷿid
                    strInsert = strInsert & mlngStockID & ","
                    '������
                    strInsert = strInsert & cboIOType.ItemData(cboIOType.ListIndex) & ","
                    '����id
                    strInsert = strInsert & lngMaterialID & ","
                    '����
                    If GetColumnPostation("����") > 0 Then
                        strTmp = IIf(IsNumeric(Trim(.TextMatrix(lngRow, .ColIndex("����")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("����"))), "")
                    Else
                        strTmp = ""
                    End If
                    dblQTY = IIf(strTmp = "", 0, strTmp)
                    strInsert = strInsert & GetFormat(dblQTY * dblTransVal, g_С��λ��.obj_���С��.����С��) & ","
                    '�ɱ���
                    If GetColumnPostation("�ɱ���") > 0 Then
                        strTmp = IIf(IsNumeric(Trim(.TextMatrix(lngRow, .ColIndex("�ɱ���")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("�ɱ���"))), "")
                    Else
                        strTmp = ""
                    End If
                    dblCostPrice = CDbl(IIf(strTmp = "", 0, strTmp))
                    strInsert = strInsert & GetFormat(dblCostPrice / IIf(dblTransVal = 0, 1, dblTransVal), g_С��λ��.obj_���С��.�ɱ���С��) & ","
                    '�ɱ����
                    If GetColumnPostation("�ɱ����") > 0 Then
                        strTmp = IIf(IsNumeric(Trim(.TextMatrix(lngRow, .ColIndex("�ɱ����")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("�ɱ����"))), "")
                    Else
                        strTmp = ""
                    End If
                    dblCost = CDbl(IIf(strTmp = "", 0, strTmp))
                    strInsert = strInsert & GetFormat(dblCost, g_С��λ��.obj_���С��.���С��) & ","
                    '�ۼ�
                    dblSalePrice = IIf(bytLotPrice = 1, dblCostPrice / IIf(dblTransVal = 0, 1, dblTransVal) * (dblAddRate + 1), dblCurSale)
                    strInsert = strInsert & GetFormat(dblSalePrice, g_С��λ��.obj_���С��.���ۼ�С��) & ","
                    '�ۼ۽��
                    dblSale = GetFormat(dblQTY * dblTransVal, g_С��λ��.obj_���С��.����С��) * GetFormat(dblSalePrice, g_С��λ��.obj_���С��.���ۼ�С��)
                    strInsert = strInsert & GetFormat(dblSale, g_С��λ��.obj_���С��.���С��) & ","
                    '���
                    strInsert = strInsert & GetFormat(dblSale, g_С��λ��.obj_���С��.���С��) - GetFormat(dblCost, g_С��λ��.obj_���С��.���С��) & ","
                    '���۲��
                    strInsert = strInsert & "null,"
                    '������
                    strInsert = strInsert & "'" & gstrUserName & "',"
                    '��������
                    strInsert = strInsert & "to_date('" & Format(Now(), "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd HH24:MI:SS'),"
                    'ժҪ
                    strInsert = strInsert & "null,"
                    '����
                    If GetColumnPostation("����") > 0 Then
                        strTmp = UCase(Trim(.TextMatrix(lngRow, .ColIndex("����"))))
                    Else
                        strTmp = ""
                    End If
                    strInsert = strInsert & IIf(strTmp = "", "null", "'" & strTmp & "'") & ","
                    '����
                    If GetColumnPostation("����") > 0 Then
                        strTmp = Trim(.TextMatrix(lngRow, .ColIndex("����")))
                    Else
                        strTmp = ""
                    End If
                    strInsert = strInsert & IIf(strTmp = "", "null", "'" & strTmp & "'") & ","
                    '��������
                    If GetColumnPostation("��������") > 0 Then
                        strTmp = IIf(IsDate(Trim(.TextMatrix(lngRow, .ColIndex("��������")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("��������"))), "")
                    Else
                        strTmp = ""
                    End If
                    If IsNumeric(strTmp) Then
                        strTmp = TranNumToDate(strTmp)
                    End If
                    strInsert = strInsert & IIf(strTmp = "", "null", "to_date('" & Format(strTmp, "yyyy-mm-dd") & "','YYYY-MM-DD')") & ","
                    'Ч��
                    If blnLot Then  '�ⷿ��������Ч��
                        If GetColumnPostation("Ч��") > 0 Then
                             strTmp = IIf(IsDate(Trim(.TextMatrix(lngRow, .ColIndex("Ч��")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("Ч��"))), "")
                        Else
                            strTmp = ""
                        End If
                        If IsNumeric(strTmp) Then
                            strTmp = TranNumToDate(strTmp)
                        End If
                        strInsert = strInsert & IIf(strTmp = "", "null", "to_date('" & Format(strTmp, "yyyy-mm-dd") & "','YYYY-MM-DD')") & ","
                    Else
                        strInsert = strInsert & "null,"
                    End If
                    If blnOnce Then
                        '�������
                        If GetColumnPostation("�������") > 0 Then
                            strTmp = IIf(IsDate(Trim(.TextMatrix(lngRow, .ColIndex("�������")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("�������"))), "")
                        Else
                            strTmp = ""
                        End If
                        If IsNumeric(strTmp) Then
                            strTmp = TranNumToDate(strTmp)
                        End If
                        strInsert = strInsert & IIf(strTmp = "", "null", "to_date('" & Format(strTmp, "yyyy-mm-dd") & "','YYYY-MM-DD')") & ","
                        '���Ч��
                        If GetColumnPostation("���Ч��") > 0 Then
                            strTmp = IIf(IsDate(Trim(.TextMatrix(lngRow, .ColIndex("���Ч��")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("���Ч��"))), "")
                        Else
                            strTmp = ""
                        End If
                        strInsert = strInsert & IIf(strTmp = "", "null )", "to_date('" & Format(strTmp, "yyyy-mm-dd") & "','YYYY-MM-DD') )")
                    Else
                        strInsert = strInsert & "null,null)"
                    End If
                End If
                Call zlDatabase.ExecuteProcedure(strInsert, Me.Caption)
            End If
        Next
    End With
    
    If blnTran = True Then
        gcnOracle.CommitTrans
    End If
    mblnResult = True
    MsgBox "����ɹ���", vbInformation, gstrSysName
    Exit Sub
    
ErrHandle:
    If blnTran = True Then
        gcnOracle.RollbackTrans
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    MsgBox "����ɹ���", vbInformation, gstrSysName
    Call SaveErrLog
End Sub

Private Function GetColIndex(ByVal strColName As String) As Long
    Dim i As Long
    For i = 1 To mobjWS.UsedRange.Columns.count
        If mobjWS.UsedRange.cells(1, i) = strColName Then
            GetColIndex = i
            Exit Function
        End If
    Next
End Function

Private Sub vsfError_EnterCell()
    Dim strTemp As String
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strCol As String
    
    With vsfError
        If .Row = 0 Then Exit Sub
        .FocusRect = flexFocusSolid
        If InStr(1, .TextMatrix(.Row, 1), "��") = 0 Then Exit Sub
        If .TextMatrix(.Row, 1) <> "" Then
            strTemp = .TextMatrix(.Row, 1)
            lngRow = Mid(strTemp, 1, InStr(1, strTemp, "��") - 1)
            strCol = Mid(strTemp, InStr(1, strTemp, "��") + 1, InStr(1, strTemp, "��") - InStr(1, strTemp, "��") - 1)
            lngCol = vsfList.ColIndex(strCol)
            If lngRow > vsfList.Rows - 1 Then MsgBox "���������Ѿ���ɾ���ˣ�", vbInformation, gstrSysName: Exit Sub
            vsfList.Row = lngRow
            vsfList.Col = lngCol
            vsfList.ShowCell lngRow, lngCol
        End If
    End With
End Sub

Private Sub vsfList_DblClick()
    With vsfList
        .EditCell
        .EditSelStart = 0
        .EditSelLength = Len(.EditText)
    End With
End Sub

Private Sub vsfList_EnterCell()
    With vsfList
        If .Row < 1 Then Exit Sub
        .FocusRect = flexFocusSolid
    End With
End Sub

Private Sub vsfList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim bln�ɱ���� As Boolean
    Dim bln��Ʊ��� As Boolean
    Dim dbl�ɱ���� As Double
    Dim dbl��Ʊ��� As Double
    Dim lngRow As Long
    
    If KeyCode = vbKeyDelete Then
        If MsgBox("��ɾ����" & vsfList.Row & "�������Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            With vsfList
                .RemoveItem .Row
                bln�ɱ���� = IIf(GetColumnPostation("�ɱ����") > 0, True, False)
                bln��Ʊ��� = IIf(GetColumnPostation("��Ʊ���") > 0, True, False)
                For lngRow = 1 To .Rows - 1
                    If bln�ɱ���� = True Then
                        dbl�ɱ���� = dbl�ɱ���� + Val(.TextMatrix(lngRow, .ColIndex("�ɱ����")))
                    End If
                    If bln��Ʊ��� = True Then
                        dbl��Ʊ��� = dbl��Ʊ��� + Val(.TextMatrix(lngRow, .ColIndex("��Ʊ���")))
                    End If
                Next
                lblCollect.Caption = "�ɱ���" & Format(dbl�ɱ����, mFMT.FM_���) & "Ԫ          ��Ʊ��" & Format(dbl��Ʊ���, mFMT.FM_���) & "Ԫ"
            End With
        End If
    End If
End Sub

Private Sub vsfList_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngRow As Long
    Dim lngCol As Long
    Dim dbl�ɱ���� As Double
    Dim dbl��Ʊ��� As Double
    Dim bln�ɱ���� As Boolean
    Dim bln��Ʊ��� As Boolean
    Dim strTemp As String
    
    Dim cbrControl As CommandBarControl
    If mbyt���뷽ʽ = 1 Then
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLSAVE, , True)
        cbrControl.Enabled = False
    End If
    
    With vsfList
        strTemp = .EditText
        If .TextMatrix(0, Col) = "�ɱ����" Then
            strTemp = Format(Val(strTemp), mFMT.FM_���)
            dbl�ɱ���� = 0
            dbl��Ʊ��� = 0
            bln��Ʊ��� = IIf(GetColumnPostation("��Ʊ���") > 0, True, False)
            For lngRow = 1 To .Rows - 1
                If lngRow = Row Then
                    dbl�ɱ���� = dbl�ɱ���� + strTemp
                Else
                    dbl�ɱ���� = dbl�ɱ���� + Val(.TextMatrix(lngRow, .ColIndex("�ɱ����")))
                End If
                If bln��Ʊ��� = True Then
                    dbl��Ʊ��� = dbl��Ʊ��� + Val(.TextMatrix(lngRow, .ColIndex("��Ʊ���")))
                End If
            Next
            .EditText = strTemp
        End If
        If .TextMatrix(0, Col) = "��Ʊ���" Then
            strTemp = Format(Val(strTemp), mFMT.FM_���)
            dbl�ɱ���� = 0
            dbl��Ʊ��� = 0
            bln�ɱ���� = IIf(GetColumnPostation("�ɱ����") > 0, True, False)
            For lngRow = 1 To .Rows - 1
                If lngRow = Row Then
                    dbl��Ʊ��� = dbl��Ʊ��� + strTemp
                Else
                    dbl��Ʊ��� = dbl��Ʊ��� + Val(.TextMatrix(lngRow, .ColIndex("��Ʊ���")))
                End If
                If bln�ɱ���� = True Then
                    dbl�ɱ���� = dbl�ɱ���� + Val(.TextMatrix(lngRow, .ColIndex("�ɱ����")))
                End If
            Next
            .EditText = strTemp
        End If
        If .TextMatrix(0, Col) = "��Ʊ���" Or .TextMatrix(0, Col) = "�ɱ����" Then
            lblCollect.Caption = "�ɱ���" & Format(dbl�ɱ����, mFMT.FM_���) & "Ԫ          ��Ʊ��" & Format(dbl��Ʊ���, mFMT.FM_���) & "Ԫ"
        End If
        If .TextMatrix(0, Col) = "����" Then
            .EditText = Format(Val(strTemp), mFMT.FM_����)
        End If
        If .TextMatrix(0, Col) = "�ɱ���" Then
            .EditText = Format(Val(strTemp), mFMT.FM_�ɱ���)
        End If
        If .TextMatrix(0, Col) = "��������" Then
            .EditText = FormatDate(strTemp)
        End If
        If .TextMatrix(0, Col) = "Ч��" Then
            .EditText = FormatDate(strTemp)
        End If
        If .TextMatrix(0, Col) = "�������" Then
            .EditText = FormatDate(strTemp)
        End If
        If .TextMatrix(0, Col) = "���Ч��" Then
            .EditText = FormatDate(strTemp)
        End If
        If .TextMatrix(0, Col) = "��Ʊ����" Then
            .EditText = FormatDate(strTemp)
        End If
        
    End With
End Sub

