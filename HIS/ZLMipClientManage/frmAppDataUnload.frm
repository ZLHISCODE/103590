VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAppDataUnload 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ϣ����ж��"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8040
   Icon            =   "frmAppDataUnload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdNext 
      Caption         =   "��һ��(&N)"
      Height          =   345
      Left            =   6765
      TabIndex        =   6
      Top             =   4725
      Width           =   1100
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   45
      ScaleHeight     =   840
      ScaleWidth      =   7995
      TabIndex        =   4
      Top             =   0
      Width           =   7995
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ϣ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   90
         TabIndex        =   5
         Top             =   225
         Width           =   1260
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   7170
         Picture         =   "frmAppDataUnload.frx":6852
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   345
      Left            =   270
      TabIndex        =   3
      Top             =   4725
      Width           =   1100
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "��һ��(&P)"
      Height          =   345
      Left            =   5610
      TabIndex        =   2
      Top             =   4725
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   1
      Top             =   870
      Width           =   8100
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   0
      TabIndex        =   0
      Top             =   4560
      Width           =   8100
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   1485
      Top             =   4515
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppDataUnload.frx":9CD4
            Key             =   "ȫѡ"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppDataUnload.frx":A26E
            Key             =   "�����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppDataUnload.frx":A808
            Key             =   "ִ����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppDataUnload.frx":ADA2
            Key             =   "��ִ��"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppDataUnload.frx":B33C
            Key             =   "ȫ��"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   3645
      Index           =   1
      Left            =   30
      ScaleHeight     =   3645
      ScaleWidth      =   7950
      TabIndex        =   14
      Top             =   900
      Width           =   7950
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   2730
         Left            =   975
         TabIndex        =   15
         Top             =   615
         Width           =   6840
         _cx             =   12065
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
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
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�빴ѡ��Ҫж����Щϵͳ����Ϣ����"
         Height          =   180
         Index           =   5
         Left            =   975
         TabIndex        =   16
         Top             =   225
         Width           =   2880
      End
      Begin VB.Image img 
         Height          =   480
         Index           =   1
         Left            =   165
         Picture         =   "frmAppDataUnload.frx":B8D6
         Top             =   165
         Width           =   480
      End
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   3645
      Index           =   2
      Left            =   30
      ScaleHeight     =   3645
      ScaleWidth      =   7950
      TabIndex        =   7
      Top             =   900
      Width           =   7950
      Begin VB.CommandButton cmdSetup 
         Caption         =   "ж��(&U)"
         Height          =   345
         Left            =   960
         TabIndex        =   8
         Top             =   3255
         Width           =   1100
      End
      Begin MSComctlLib.ProgressBar pgb 
         Height          =   225
         Left            =   2130
         TabIndex        =   9
         Top             =   3375
         Visible         =   0   'False
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfStep 
         Height          =   2490
         Left            =   975
         TabIndex        =   10
         Top             =   600
         Width           =   6840
         _cx             =   12065
         _cy             =   4392
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
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
         FloodColor      =   192
         SheetBorder     =   -2147483643
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
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
      End
      Begin VB.Image img 
         Height          =   480
         Index           =   3
         Left            =   195
         Picture         =   "frmAppDataUnload.frx":D258
         Top             =   195
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ж�ء�����ʼж���ѹ�ѡ����Ϣ����"
         Height          =   180
         Index           =   1
         Left            =   960
         TabIndex        =   13
         Top             =   165
         Width           =   3420
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         Height          =   180
         Index           =   6
         Left            =   7395
         TabIndex        =   12
         Top             =   3405
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����ж��.."
         Height          =   180
         Index           =   12
         Left            =   2145
         TabIndex        =   11
         Top             =   3150
         Visible         =   0   'False
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmAppDataUnload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOK As Boolean
Private mobjFso As New FileSystemObject
Private mclsOracle As clsDataOracle
Private mblnStep(1 To 2) As Boolean
Private mstrManageVersion As String
Private mstrVersion As String
Private mintPage As Integer
Private mclsVsf As zlVSFlexGrid.clsVsf
Private mclsVsfStep As zlVSFlexGrid.clsVsf
Private mclsVsfUser As zlVSFlexGrid.clsVsf
Private mbytMode As Byte

Private WithEvents mobjScript As clsOracleScript
Attribute mobjScript.VB_VarHelpID = -1

Public Function ShowDialog() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    
    mblnOK = False
    
    Set mclsOracle = New clsDataOracle
    
    
    For intLoop = 1 To picPage.UBound
        picPage(intLoop).Left = 0
        picPage(intLoop).Top = 915
        picPage(intLoop).Width = 7950
        picPage(intLoop).Height = 3645
    Next
    
    Call InitGrid
    
    mbytMode = 2
    mintPage = 1
    Call ShowPage(mintPage)
    
    Me.Show 1
    ShowDialog = mblnOK
End Function

Private Function InitGrid() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Dim rsData As ADODB.Recordset
    Dim intRow As Integer
    Dim intCount As Integer
    Dim strSQL As String
    
    Set mclsVsf = New zlVSFlexGrid.clsVsf
    With mclsVsf
        Call .Initialize(Me.Controls, vsf, True, False)
        Call .ClearColumn
        
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTBoolean, "", "[ѡ��]", False)
        Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, , "Data_DB", True)
        Call .AppendColumn("code", 0, flexAlignLeftCenter, flexDTString, , "Data_Code", True)
        Call .AppendColumn("��Ϣ����", 3000, flexAlignLeftCenter, flexDTString, , "Data_Title", True)
        Call .AppendColumn("��װʱ��", 900, flexAlignLeftCenter, flexDTString, , "Setup_Time", True)
        
        Call .InitializeEdit(True, False, False)
        Call .InitializeEditColumn(vsf.ColIndex("ѡ��"), True, vbVsfEditCheck)

        vsf.Cell(flexcpPicture, 0, .ColIndex("ѡ��")) = imgList.ListImages("ȫ��").Picture
        .AppendRows = True
        
    End With
                    
    strSQL = "Select Data_Code,Data_Title,Data_Owner,Data_System,Data_DB,Setup_Time From zlMip_Data_Setup"
    Set rsData = zlDataBase.OpenSQLRecord(strSQL, gstrSysName)
    If rsData.BOF = False Then
        Call mclsVsf.LoadGrid(rsData)
    End If
                    
    'ʾ��
'    With vsf
'        .Rows = 7
'        .TextMatrix(1, 2) = "ҽԺ��Ϣ��׼��"
'        .TextMatrix(1, 3) = "2014-04-08 15:34:00"
'
'        .TextMatrix(2, 2) = "������ϵͳ"
'        .TextMatrix(2, 3) = "2014-04-08 15:34:00"
'
'        .TextMatrix(3, 2) = "�������ϵͳ"
'        .TextMatrix(3, 3) = "2014-04-08 15:34:00"
'
'        .TextMatrix(4, 2) = "Ѫ�����ϵͳ"
'        .TextMatrix(4, 3) = "2014-04-08 15:34:00"
'
'        .TextMatrix(5, 2) = "��������ϵͳ"
'        .TextMatrix(5, 3) = "2014-04-08 15:34:00"
'
'        .TextMatrix(6, 2) = "�������ϵͳ"
'        .TextMatrix(6, 3) = "2014-04-08 15:34:00"
'
'        mclsVsf.AppendRows = True
'    End With
        
    '------------------------------------------------------------------------------------------------------------------
            
    Set mclsVsfStep = New zlVSFlexGrid.clsVsf
    With mclsVsfStep
        Call .Initialize(Me.Controls, vsfStep, True, False)
        Call .ClearColumn
        
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
        Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, , "Data_DB", True)
        Call .AppendColumn("step", 1500, flexAlignLeftCenter, flexDTString, , "item_note", True)
        vsfStep.RowHidden(0) = True
    End With

    InitGrid = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Sub ShowPage(ByVal intPage As Integer)
    Dim intLoop As Integer
    
    For intLoop = 1 To picPage.UBound
        picPage(intLoop).Visible = False
    Next
    
    picPage(intPage).Visible = True
        
    cmdNext.Enabled = (intPage < picPage.UBound)
    cmdPrevious.Enabled = (intPage > 1)
    
End Sub

Private Sub SelectedAll()
    '******************************************************************************************************************
    '���ܣ����ȫѡ��ȫ��
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intRow As Integer
    
    With vsf
        Select Case mbytMode
        Case 1
            '��״̬Ϊȫѡ������Ϊȫ��
            For intRow = 1 To .Rows - 1
                .TextMatrix(intRow, .ColIndex("ѡ��")) = 0
            Next
            .Cell(flexcpPicture, 0, .ColIndex("ѡ��")) = imgList.ListImages("ȫ��").Picture
            mbytMode = 2
        Case 2
            '��״̬Ϊȫ�壬����Ϊȫѡ
            For intRow = 1 To .Rows - 1
                .TextMatrix(intRow, .ColIndex("ѡ��")) = 1
            Next
            .Cell(flexcpPicture, 0, .ColIndex("ѡ��")) = imgList.ListImages("ȫѡ").Picture
            mbytMode = 1
        End Select
    End With
    
End Sub

Private Function OpenDialog(ByRef objDlg As Object, ByVal strDialogTitle As String, ByVal strFilter As String) As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************

    Dim strTmp As String
            
    With objDlg
        .DialogTitle = strDialogTitle
        .Filter = strFilter
    
        On Error Resume Next
    
        .flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
        .FileName = ""
        .MaxFileSize = 32767
        .CancelError = True
        .ShowOpen
    
        If Err.Number = 0 And .FileName <> "" Then
    
            strTmp = .FileName
    
            On Error GoTo errHand
                                                    
            OpenDialog = strTmp
            
        Else
            Err.Clear
        End If
    End With
    
    Exit Function

errHand:
    MsgBox "���ܴ��ļ�(" & strTmp & "),���ļ���������ʹ�û��ļ�������!"
End Function

Private Function SaveDialog(ByRef objDlg As Object, ByVal strDialogTitle As String, ByVal strFilter As String) As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************

    Dim strTmp As String
            
    With objDlg
        .DialogTitle = strDialogTitle
        .Filter = strFilter
    
        On Error Resume Next
    
        .flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
        .FileName = ""
        .MaxFileSize = 32767
        .CancelError = True
        .ShowSave
    
        If Err.Number = 0 And .FileName <> "" Then
    
            strTmp = .FileName
    
            On Error GoTo errHand
                                                    
            SaveDialog = strTmp
            
        Else
            Err.Clear
        End If
    End With
    
    Exit Function

errHand:
    MsgBox "���ܱ���Ϊ�ļ�(" & strTmp & "),���ļ���������ʹ�û��ļ��Ѿ�����!"
End Function

Private Function CheckPassword(ByVal strUser As String, ByVal strPassword As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    CheckPassword = mclsOracle.OraDataOpen(gstrServerName, strUser, strPassword, True)
End Function

Private Function CheckSetupFile(ByVal strFile As String) As Boolean
    '******************************************************************************************************************
    '���ܣ������Ͱ�װ�����ļ�����ȷ��
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strIniPath As String
    Dim strTemp As String
    Dim objText As TextStream
    Dim strManageVersion As String
    Dim intLoop As Integer
    Dim aryTemp As Variant
    Dim aryItem As Variant
    
    strIniPath = Mid(strFile, 1, Len(strFile) - 11)
    
    '����ļ�ƥ���Լ��
    '------------------------------------------------------------------------------------------------------------------
    strTemp = ""
    If Dir(strIniPath & "zlMipClientStruct.SQL") = "" Then strTemp = strTemp & vbCr & "�ṹ�ļ�" & strIniPath & "zlMipClientStruct.SQL"
    If Dir(strIniPath & "zlMipClientData.SQL") = "" Then strTemp = strTemp & vbCr & "�����ļ�" & strIniPath & "zlMipClientData.SQL"
    
    If strTemp <> "" Then
        MsgBox "���°�װ������ļ���ʧ�����ܼ�����������" & strTemp, vbExclamation, gstrSysName
        Exit Function
    End If
    
    '��װ�����ļ�����
    '------------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error Resume Next
    Set objText = mobjFso.OpenTextFile(strFile)
    
    strTemp = Trim(objText.ReadLine)
    
    mstrVersion = ""
    mstrManageVersion = ""
    
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[�汾��]" Then
        mstrVersion = Trim(Mid(strTemp, 6))
    Else
        Err.Raise 10
    End If
    
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 6) = "[��Ϣ����]" Then
        strTemp = Trim(Mid(strTemp, 7))
'
'        lst.Clear
'        aryTemp = Split(strTemp, "|")
'        For intLoop = 0 To UBound(aryTemp)
'            aryItem = Split(aryTemp(intLoop), "=")
'            lst.AddItem aryItem(0)
'            lst.ItemData(lst.NewIndex) = aryItem(1)
'        Next
        
    Else
        Err.Raise 10
    End If
    
    lbl(2).Caption = "�汾�ţ�" & mstrVersion
        
    objText.Close
    
    
    CheckSetupFile = True
End Function

Public Function GetVerDouble(ByVal varVer As Variant) As Double
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    '���ܣ����ݰ汾�ַ������������ֻ��İ汾
    '������varVer   �汾�ַ�������9.5.0
    Dim varArray As Variant
    
    varVer = IIf(IsNull(varVer), "", varVer)
    varArray = Split(varVer, ".")
    
    If UBound(varArray) < 2 Then Exit Function
    
    GetVerDouble = Val(varArray(0)) * 10 ^ 8 + Val(varArray(1)) * 10 ^ 4 + Val(varArray(2))
End Function

Private Function UnloadMipData(ByVal strCode As String) As Boolean
    '******************************************************************************************************************
    '���ܣ����ݹؼ���ж�ض�Ӧϵͳ����Ϣ����
    '������strCode �ؼ���,��HIS
    '���أ�
    '******************************************************************************************************************
    Dim strSQL As String
    
    pgb.Value = 0
    lbl(6).Caption = "0%"
    'ɾ����װ��¼��
    strSQL = "Delete From zlmip_data_setup Where data_code='" & strCode & "'"
    gclsMsgOracle.ExecuteSQL strSQL, gstrSysName
    
    pgb.Value = 25
    lbl(6).Caption = "25%"
    'ɾ�����¼
    '--zlmip_tabext_condition
    strSQL = "Delete From zlmip_tabext_condition Where ext_id In(Select ID From zlmip_tab_extend Where source_tab_id In(Select id From zlmip_table Where data_code='" & strCode & "'))"
    gclsMsgOracle.ExecuteSQL strSQL, gstrSysName
    '--zlmip_tab_extend
    strSQL = "Delete From zlmip_tab_extend Where source_tab_id In(Select id From zlmip_table Where data_code='" & strCode & "')"
    gclsMsgOracle.ExecuteSQL strSQL, gstrSysName
    '--zlmip_tab_parameter
    strSQL = "Delete From zlmip_tab_parameter Where tab_id In(Select id From zlmip_table Where data_code='" & strCode & "')"
    gclsMsgOracle.ExecuteSQL strSQL, gstrSysName
    '--zlmip_tab_field
    strSQL = "Delete From zlmip_tab_field Where tab_id In(Select id From zlmip_table Where data_code='" & strCode & "')"
    gclsMsgOracle.ExecuteSQL strSQL, gstrSysName
    '--zlmip_table
    strSQL = "Delete From zlmip_table Where data_code='" & strCode & "'"
    gclsMsgOracle.ExecuteSQL strSQL, gstrSysName
    
    'ɾ��Send_Log��ؼ�¼
    '--zlmip_sendlog_again
    strSQL = "Delete From zlmip_sendlog_again Where send_log_id In(Select id From zlmip_send_log Where item_id In(Select id From zlmip_item Where data_code='" & strCode & "'))"
    gclsMsgOracle.ExecuteSQL strSQL, gstrSysName
    '--zlmip_sendlog_parameter
    strSQL = "Delete From zlmip_sendlog_parameter Where send_log_id In(Select id From zlmip_send_log Where item_id In(Select id From zlmip_item Where data_code='" & strCode & "'))"
    gclsMsgOracle.ExecuteSQL strSQL, gstrSysName
    '--zlmip_send_log
    strSQL = "Delete From zlmip_send_log Where item_id In(Select id From zlmip_item Where data_code='" & strCode & "')"
    gclsMsgOracle.ExecuteSQL strSQL, gstrSysName
    
    pgb.Value = 50
    lbl(6).Caption = "50%"
    'ɾ��item��¼
    '--zlmip_item_deliver
    strSQL = "Delete From zlmip_item_deliver Where item_id In(Select id From zlmip_item Where data_code='" & strCode & "')"
    gclsMsgOracle.ExecuteSQL strSQL, gstrSysName
    '--zlmip_item_config
    strSQL = "Delete From zlmip_item_config Where item_id In(Select id From zlmip_item Where data_code='" & strCode & "')"
    gclsMsgOracle.ExecuteSQL strSQL, gstrSysName
    '--zlmip_item_frequency
    strSQL = "Delete From zlmip_item_frequency Where item_id In(Select id From zlmip_item Where data_code='" & strCode & "')"
    gclsMsgOracle.ExecuteSQL strSQL, gstrSysName
    '--zlmip_item_field
    strSQL = "Delete From zlmip_item_frequency Where item_id In(Select id From zlmip_item Where data_code='" & strCode & "')"
    gclsMsgOracle.ExecuteSQL strSQL, gstrSysName
    '--zlmip_item_parameter
    strSQL = "Delete From zlmip_item_parameter Where item_id In(Select id From zlmip_item Where data_code='" & strCode & "')"
    gclsMsgOracle.ExecuteSQL strSQL, gstrSysName
    '--zlmip_item
    strSQL = "Delete From zlmip_item Where data_code='" & strCode & "'"
    gclsMsgOracle.ExecuteSQL strSQL, gstrSysName
    
    pgb.Value = 100
    lbl(6).Caption = "100%"
    
    UnloadMipData = True
End Function

Public Function SetupMipClient(ByVal strInstallFile As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strPath As String
    Dim intLoop As Integer
    Dim strSQL As String
    Dim intPercent As Integer
    
    On Error GoTo errHand
    
    strPath = Left(strInstallFile, Len(strInstallFile) - Len("zlSetup.ini"))
    
    '��װ�ṹ
    '------------------------------------------------------------------------------------------------------------------
    Set mobjScript = New clsOracleScript
    If mobjScript.OpenScriptFile(strPath & "zlMipClientStruct.SQL") Then
        
        lbl(4).Caption = "����ִ�нṹ�ű�..."
'        pgb.Value = 0
    
        For intLoop = 1 To mobjScript.SQLCount
            Call mobjScript.ExecuteSQL(mclsOracle.DatabaseConnection, mobjScript.SQL(intLoop))
'            intPercent = 100 * intLoop / mobjScript.SQLCount
'            If pgb.Value <> intPercent Then pgb.Value = intPercent
        Next
    End If
    
    '��װ����
    '------------------------------------------------------------------------------------------------------------------
    If mobjScript.OpenScriptFile(strPath & "zlMipClientData.SQL") Then
        lbl(4).Caption = "����ִ�����ݽű�..."
'        pgb.Value = 0
        For intLoop = 1 To mobjScript.SQLCount
            Call mobjScript.ExecuteSQL(mclsOracle.DatabaseConnection, mobjScript.SQL(intLoop))
'            intPercent = 100 * intLoop / mobjScript.SQLCount
'            If pgb.Value <> intPercent Then pgb.Value = intPercent
        Next
    End If
    
    '��¼��װ�汾
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "Insert Into zlRegInfo(��Ŀ,�к�,����) Select '��Ϣ����ƽ̨�ͻ���',1,'" & mstrVersion & "' From Dual"
    mclsOracle.ExecuteSQL strSQL, gstrSysName
    
    SetupMipClient = True
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If MsgBox("�������д����Ƿ������" & vbCrLf & "    " & Err.Description, vbQuestion Or vbYesNo, gstrSysName) = vbYes Then
        Resume
    End If
    
    'ж���Ѿ���װ������
    '------------------------------------------------------------------------------------------------------------------
'    lbl(4).Caption = "����ж���Ѿ���װ������..."
    
End Function


Private Sub cmdNext_Click()
    
    Dim intCount As Integer
    Dim intRow As Integer
    
    Select Case mintPage
    '------------------------------------------------------------------------------------------------------------------
    Case 1
        With vsfStep
            .Rows = 0
            intCount = 0
            For intRow = 1 To vsf.Rows - 1
                If Abs(Val(vsf.TextMatrix(intRow, vsf.ColIndex("ѡ��")))) = 1 Then
                    .Rows = .Rows + 1
                    .TextMatrix(intCount, .ColIndex("ID")) = vsf.TextMatrix(intRow, vsf.ColIndex("ID"))
                    .TextMatrix(intCount, .ColIndex("step")) = "ж��" & vsf.TextMatrix(intRow, vsf.ColIndex("��Ϣ����")) & "��Ϣ����"
                    .Cell(flexcpPicture, intCount, .ColIndex("ͼ��"), intCount, .ColIndex("ͼ��")) = imgList.ListImages("��ִ��").Picture
                    intCount = intCount + 1
                End If
            Next
        End With
        mintPage = mintPage + 1
        Call ShowPage(mintPage)
    End Select
    
    
End Sub

Private Sub cmdPrevious_Click()

    Select Case mintPage
    '------------------------------------------------------------------------------------------------------------------
    Case 2
        
        mintPage = mintPage - 1
        Call ShowPage(mintPage)
            
    End Select
    
End Sub

Private Sub cmdSetup_Click()
    Dim intRow As Integer
    Dim intCount As Integer
    
    If MsgBox("ȷ����Ҫж����Ϣ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    With vsf
        lbl(12).Visible = True
        lbl(6).Visible = True
        pgb.Visible = True
        intCount = 0
        For intRow = 1 To .Rows - 1
            If Abs(Val(.TextMatrix(intRow, .ColIndex("ѡ��")))) = 1 Then
                lbl(12).Caption = "����ж��" & .TextMatrix(intRow, .ColIndex("��Ϣ����")) & "��Ϣ����"
                vsfStep.Cell(flexcpPicture, intCount, vsfStep.ColIndex("ͼ��")) = imgList.ListImages("ִ����").Picture
                Call UnloadMipData(.TextMatrix(intRow, .ColIndex("code")))
                vsfStep.Cell(flexcpPicture, intCount, vsfStep.ColIndex("ͼ��")) = imgList.ListImages("�����").Picture
                intCount = intCount + 1
            End If
        Next
    End With
    lbl(12).Caption = "����ж�سɹ�!"
    MsgBox "ж�سɹ�!", vbInformation + vbOKOnly, gstrSysName
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    If Not (mclsOracle Is Nothing) Then
'        Set mclsOracle = Nothing
'    End If
'
'    Dim frmThis As Form
'
'    On Error Resume Next
'
'    '�رձ���������
'    For Each frmThis In Forms
'        If frmThis.Caption <> Me.Caption Then
'            Unload frmThis
'        End If
'    Next
'
End Sub

Private Sub mobjScript_AfterAnalyseLine(ByVal Line As Long, ByVal Lines As Long)
    Dim intPercent As Integer
    
'    If pgb.Visible = False Then pgb.Visible = True
'    If lbl(4).Visible = False Then
'        lbl(4).Visible = True
'        lbl(4).Caption = "���ڷ����ű��ļ�...."
'    End If
'
'    intPercent = 100 * Line / Lines
'    If pgb.Value <> intPercent Then pgb.Value = intPercent
'
End Sub

Private Sub mobjScript_BeforeAnalyseLine(ByVal Line As Long, ByVal Lines As Long)
'    If pgb.Visible = False Then pgb.Visible = True
'    If lbl(4).Visible = False Then
'        lbl(4).Visible = True
'        lbl(4).Caption = "���ڷ����ű��ļ�...."
'    End If
End Sub


Private Sub vsf_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_Click()
    If vsf.MouseRow = 0 And vsf.Col = vsf.ColIndex("ѡ��") Then
        Call SelectedAll
    End If
End Sub

