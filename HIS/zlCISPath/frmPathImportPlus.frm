VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPathImportPlus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ٴ�·��ѡ��"
   ClientHeight    =   9960
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12255
   Icon            =   "frmPathImportPlus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9960
   ScaleWidth      =   12255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8370
      Left            =   0
      ScaleHeight     =   8370
      ScaleWidth      =   12255
      TabIndex        =   5
      Top             =   800
      Width           =   12255
      Begin VSFlex8Ctl.VSFlexGrid vsPath 
         Height          =   4065
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   12015
         _cx             =   1973310153
         _cy             =   1973296130
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
         BackColorFixed  =   15597549
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16444122
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   32768
         GridColorFixed  =   32768
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   320
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPathImportPlus.frx":6852
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
         BackColorFrozen =   14811105
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vsDisease 
         Height          =   3465
         Left            =   120
         TabIndex        =   14
         Top             =   4800
         Width           =   12015
         _cx             =   1973310153
         _cy             =   1973295072
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
         BackColorFixed  =   15597549
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16444122
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   32768
         GridColorFixed  =   32768
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   320
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPathImportPlus.frx":68CB
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
         BackColorFrozen =   14811105
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����±���ѡ��һ�������ڸò��˵����"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   4560
         Width           =   3240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   120
         X2              =   12840
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   120
         X2              =   12840
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����±���ѡ��һ�������ڸò��˵��ٴ�·��"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   3600
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   12255
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   9165
      Width           =   12255
      Begin VB.CommandButton cmdPathOut 
         Caption         =   "��������"
         Height          =   350
         Left            =   9720
         TabIndex        =   15
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdPathIn 
         Caption         =   "�뾶����"
         Height          =   350
         Left            =   10920
         TabIndex        =   8
         Top             =   240
         Width           =   1100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   6
         X1              =   0
         X2              =   12720
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   7
         X1              =   0
         X2              =   12720
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   12255
      TabIndex        =   1
      Top             =   0
      Width           =   12255
      Begin VB.Frame fraSel 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFF0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   240
         TabIndex        =   10
         Top             =   277
         Width           =   3495
         Begin VB.OptionButton optSel 
            Appearance      =   0  'Flat
            BackColor       =   &H00EFF0E0&
            Caption         =   "����ϱ�������"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton optSel 
            BackColor       =   &H00EFF0E0&
            Caption         =   "��������������"
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   11
            Top             =   0
            Width           =   1575
         End
      End
      Begin VB.Frame fraDiag 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFF0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   3840
         TabIndex        =   2
         Top             =   230
         Width           =   5055
         Begin VB.CommandButton cmd 
            Caption         =   "��"
            Height          =   285
            Left            =   3660
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "ѡ�����"
            Top             =   35
            Width           =   285
         End
         Begin VB.TextBox txtDiagnose 
            Height          =   330
            Left            =   480
            TabIndex        =   3
            ToolTipText     =   "¼����ϲ���·��"
            Top             =   0
            Width           =   3495
         End
         Begin VB.Label lblDiagnose 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���"
            Height          =   180
            Left            =   0
            TabIndex        =   4
            Top             =   75
            Width           =   360
         End
      End
      Begin MSComctlLib.ImageList imgSrc 
         Left            =   11520
         Top             =   120
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
               Picture         =   "frmPathImportPlus.frx":692B
               Key             =   "chkRed"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPathImportPlus.frx":6EC5
               Key             =   "unchkRed"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPathImportPlus.frx":745F
               Key             =   "chkRedUnSquare"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPathImportPlus.frx":79F9
               Key             =   "unchkBlue"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPathImportPlus.frx":7F93
               Key             =   "chkBlue"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmPathImportPlus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mPati As TYPE_Pati
Private mPP As TYPE_PATH_Pati

Private mfrmParent As Object

Private mstrSex    As String
Private mblnOK As Boolean
Private mbln��ҽ�� As Boolean
Private mbln��� As Boolean

Private mintDiagInput As Integer        'mintDiagInput:1-��ҽ��ѡ��������Դ,2-������ϱ�׼����,3-���ռ�����������
Private mintDiagInputZY As Integer      'ϵͳ����ѡ��������뷽ʽ��mintDiagInput=1��ʱסԺ��ҽ��ϣ�0-������ϱ�׼����,1-���ݼ�����������
Private mintDiagInputXY As Integer      'ϵͳ����ѡ��������뷽ʽ��mintDiagInput=1��ʱסԺ��ҽ��ϣ�0-������ϱ�׼����,1-���ݼ�����������
Private mintDiag As Integer             '��¼��ϱ������뷽ʽ

Private mrsPati As ADODB.Recordset
Private mrsPath As ADODB.Recordset      '����·��
Private mrsPathDept As ADODB.Recordset      '����·��
Private mrsDisease As ADODB.Recordset   '���没��

Private mcolPati As Collection

Private mblnICD11 As Boolean
Private mblnHave As Boolean    'T-�������

Private Enum E_���
    E_IX_����� = 0
    E_IX_������ = 1
End Enum

Public Function ShowMe(frmParent As Object, t_pati As TYPE_Pati, ByRef t_pp As TYPE_PATH_Pati) As Boolean
'������
    mPati = t_pati
    mPP = t_pp
    Set mfrmParent = frmParent
    mbln��ҽ�� = Sys.DeptHaveProperty(mPati.����ID, "��ҽ��")
    
    Set mrsPati = GetPatiInfo(mPati.����ID, mPati.��ҳID, mcolPati)
    If mrsPati.RecordCount = 0 Then
        MsgBox "��ȡ���˵�ǰסԺ��Ϣʧ�ܡ�", vbInformation, gstrSysName
        Exit Function
    End If
    Set mrsPathDept = GetPathTable(0, 0, mPati.����ID, -1)
    If mrsPathDept.RecordCount = 0 Then
        MsgBox "������û�з��ϵ�ǰ���˵���Ч�ٴ�·����", vbInformation, gstrSysName
        Exit Function
    End If
    mblnICD11 = IsICDElevent()
    mblnOK = False
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Sub cmdPathOut_Click()
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Load()
    Dim bln�������� As String
    '������ȡ
    
    bln�������� = InStr(GetInsidePrivs(pסԺҽ��վ), "��������") > 0
    mintDiagInput = Val(zlDatabase.GetPara(55, glngSys, , 1))
    mintDiagInputXY = Val(zlDatabase.GetPara("��ҽ�������", glngSys, pסԺҽ��վ, 0, Array(optSel(E_IX_������), optSel(E_IX_�����)), bln��������))
    If mbln��ҽ�� Then
        mintDiagInputZY = Val(zlDatabase.GetPara("��ҽ�������", glngSys, pסԺҽ��վ, 0, Array(optSel(E_IX_������), optSel(E_IX_�����)), bln��������))
    End If

    '����·��, a.�����Ա�, a.��������, a.���°汾,Nvl(a.��������,'��') as ��������
    Call Grid.Init(vsPath, "ѡ��,600,4;����,1200,1;����,3495,1;˵��,4995,1")
    Call Grid.Init(vsDisease, "ѡ��,600,4;����,1200,1;����,3495,1;���")
    With vsPath
        .RowHeightMin = 330
    End With
    With vsDisease
        .RowHeightMin = 330
    End With
    Set mrsPath = mrsPathDept
    Call LoadPath(mrsPath)
    If mbln��� Then mbln��� = False: Exit Sub
End Sub

Private Sub cmdPathIn_Click()
    '���ݵ���·��
    If Not SaveData() Then Exit Sub
    mblnOK = True
End Sub

Private Sub Form_Activate()
    Call txtDiagnose.SetFocus
End Sub

Private Sub ResizeDiagWay()
'����:�������¼�뷽ʽ
    If mblnICD11 Then
        fraSel.Visible = False
        fraDiag.Left = fraSel.Left
        mintDiag = 1
    Else
        'ICD-10
        If mintDiagInput = 1 Then
            fraSel.Visible = True
            If mbln��ҽ�� Then
                mintDiag = mintDiagInputZY
            Else
                mintDiag = mintDiagInputXY
            End If
        Else
            fraSel.Visible = False
            fraDiag.Left = fraSel.Left
            mintDiag = mintDiagInput - 2
        End If
    End If
    optSel(mintDiag).Value = True
End Sub

Private Sub LoadPath(ByVal rsPath As ADODB.Recordset)
    Dim i As Long
    
    rsPath.Filter = ""
    vsDisease.Rows = 1: vsDisease.Rows = 2
    With vsPath
        .Rows = 1 '�����ʷ����
        .Rows = rsPath.RecordCount + 1
        .AllowUserResizing = flexResizeColumns
        If .Rows = 1 Then .Rows = .Rows + 1
        For i = 1 To rsPath.RecordCount
            .RowData(i) = Val(rsPath!ID & "")
            .Cell(flexcpData, i, .ColIndex("����")) = Val(rsPath!���°汾 & "")
            .Cell(flexcpPictureAlignment, i, .ColIndex("ѡ��")) = flexAlignCenterCenter
            .TextMatrix(i, .ColIndex("����")) = rsPath!���� & ""
            .TextMatrix(i, .ColIndex("����")) = rsPath!���� & ""
            .TextMatrix(i, .ColIndex("˵��")) = rsPath!˵�� & ""
            rsPath.MoveNext
        Next
        If rsPath.RecordCount = 1 Then
            .Row = 1
            Set .Cell(flexcpPicture, .Row, .ColIndex("ѡ��")) = imgSrc.ListImages("chkRedUnSquare").Picture
            Call LoadDisease(.RowData(.Row)) '��һ��ʱĬ�ϼ��ز���
        End If
        cmdPathIn.Enabled = rsPath.RecordCount > 0
    End With
End Sub

Private Sub LoadDisease(ByVal lngPathID As Long)
    Dim i As Long, lngSel As Long
    Dim rsTmp As ADODB.Recordset
    Dim blnRead As Boolean
    
    If mrsDisease Is Nothing Then
        Call InitRSDisease
        blnRead = True
    Else
        mrsDisease.Filter = "·��ID = " & lngPathID
        If mrsDisease.RecordCount = 0 Then blnRead = True
    End If
    If blnRead = True Then
        Set rsTmp = GetPathDisease(lngPathID)
        Do While Not rsTmp.EOF
            '·��ID,����ID,������,������,���ID,�����,�����
            mrsDisease.AddNew Array("·��ID", "����ID", "������", "������", "���ID", "�����", "�����", "���"), _
            Array(lngPathID, Nvl(rsTmp!����id, 0), rsTmp!������ & "", rsTmp!������ & "", Nvl(rsTmp!���id, 0), rsTmp!����� & "", rsTmp!����� & "", rsTmp!��� & "")
            rsTmp.MoveNext
        Loop
        mrsDisease.Filter = "·��ID = " & lngPathID
    End If
    
    With vsDisease
        .Rows = 1: '�����ʷ����
        .Rows = mrsDisease.RecordCount + 1
        If .Rows = 1 Then .Rows = .Rows + 1
        For i = 1 To mrsDisease.RecordCount
            .RowData(i) = Val(mrsDisease!����id & "")
            .Cell(flexcpPictureAlignment, i, .ColIndex("ѡ��")) = flexAlignCenterCenter
            .Cell(flexcpData, i, .ColIndex("����")) = Val(mrsDisease!���id & "")
            .TextMatrix(i, .ColIndex("����")) = IIf(Val(mrsDisease!����id & "") > 0, mrsDisease!������ & "", mrsDisease!����� & "")
            .TextMatrix(i, .ColIndex("����")) = IIf(Val(mrsDisease!����id & "") > 0, mrsDisease!������ & "", mrsDisease!����� & "")
            .TextMatrix(i, .ColIndex("���")) = mrsDisease!��� & ""
            If mrsDisease!����id & "" = txtDiagnose.Tag Or mrsDisease!���id & "" = txtDiagnose.Tag Then
                .Row = i 'ȱʡ����λ���� �������ҽ���һ��ȱʡ��λ;���ֺ�¼�����һ��ȱʡ��λ
            End If
            mrsDisease.MoveNext
        Next
        If mrsDisease.RecordCount > 0 And .Row < 1 Then
            .Row = 1
            Set .Cell(flexcpPicture, .Row, .ColIndex("ѡ��")) = imgSrc.ListImages("chkRedUnSquare").Picture
        End If
        If .Row > 0 Then
            .ShowCell .Row, .ColIndex("����")
        End If
    End With
End Sub

Private Sub InitRSDisease()
'����:��ʼ����¼
'     ����: ·��ID,����ID,������,������,���ID,�����,�����
    Set mrsDisease = New ADODB.Recordset
    mrsDisease.Fields.Append "·��ID", adBigInt
    mrsDisease.Fields.Append "����ID", adBigInt
    mrsDisease.Fields.Append "������", adVarChar, 20, adFldIsNullable
    mrsDisease.Fields.Append "������", adVarChar, 200, adFldIsNullable
    mrsDisease.Fields.Append "���ID", adBigInt
    mrsDisease.Fields.Append "�����", adVarChar, 20, adFldIsNullable
    mrsDisease.Fields.Append "�����", adVarChar, 200, adFldIsNullable
    mrsDisease.Fields.Append "���", adVarChar, 1, adFldIsNullable

    mrsDisease.CursorLocation = adUseClient
    mrsDisease.LockType = adLockOptimistic
    mrsDisease.CursorType = adOpenStatic
    mrsDisease.Open
End Sub

Private Sub Form_Resize()
    Call ResizeDiagWay
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Not mblnOK Then
        Cancel = 1 'ֻ��ͨ����ťȡ��
    Else
        Set mrsPath = Nothing
        Set mrsDisease = Nothing
        Set mrsPathDept = Nothing
        Set mrsPati = Nothing
        Set mcolPati = Nothing
    End If
End Sub

Private Sub optSel_Click(Index As Integer)
    If optSel(Index).Value Then mintDiag = Index
End Sub

Private Sub txtDiagnose_Change()
    If Trim(txtDiagnose.Text) = "" And txtDiagnose.Tag <> "" Then
        txtDiagnose.Tag = ""
        Set mrsPath = mrsPathDept
        Call LoadPath(mrsPath)
    End If
End Sub

Private Sub txtDiagnose_GotFocus()
    Call zlControl.TxtSelAll(txtDiagnose)
End Sub

Private Sub txtDiagnose_KeyPress(KeyAscii As Integer)
    Dim strInput As String
    Dim strSql As String
    Dim vRect As RECT
    Dim blnCancel As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim lngNum As Long
    Dim strType As String
    
    If KeyAscii = 13 Then
        strInput = UCase(Trim(txtDiagnose.Text))
        txtDiagnose.Tag = ""
        If strInput = "" Then
'            Set mrsPath = mrsPathDept
'            Call LoadPath(mrsPath)
            Exit Sub
        End If
        If mblnICD11 Then
            If mbln��ҽ�� Then
                strType = ",1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,24,25,26,27,"
                strSql = "Select ��� From ����������� Where �½� = [1] And ���� = [2] And ���� = [3]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, "26", "��ͳҽѧ֤��TM1��", "L1-SE7")
                If Not rsTmp.EOF Then
                    lngNum = Val("" & rsTmp!���)
                    strSql = IIf(lngNum <> 0, " And a.����ID Not In (Select e.ID From ����������� e where e.�½�='26' And e.���>=" & lngNum & ")", "")
                End If
            Else
                strType = ",1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,24,25,27,"
            End If

            If zlCommFun.IsCharChinese(strInput) Then
                strSql = strSql & " And A.���� Like [2]" '���뺺��ʱ,ֻƥ������
            Else
                strSql = strSql & " And (A.���� Like [1] Or A.���� Like [2] Or " & IIf(gint���� = 0, "A.����", "A.�����") & " Like [2])"
            End If

            strSql = _
                " Select A.ID,A.ID as ��ĿID,A.����,A.����,A.����," & IIf(gint���� = 0, "A.����", "A.����� as ����") & ",A.˵��" & _
                " From ��������Ŀ¼ A Where A.��� ='E' And Instr([5],','||A.�½�||',')>0 " & strSql & _
                IIf(mstrSex <> "", " And (A.�Ա�����=[3] Or A.�Ա����� is NULL)", "") & _
                " And (A.����ʱ�� is Null Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Order by A.����"
        Else
            If mintDiag = E_IX_����� Then
                '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
                If zlCommFun.IsCharChinese(strInput) Then
                    strSql = "B.���� Like [2]" '���뺺��ʱ,ֻƥ������
                Else
                    strSql = "A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]"
                End If
                strSql = _
                    " Select Distinct A.ID,A.ID as ��ĿID,A.����,A.����,A.˵��,A.����" & _
                    " From �������Ŀ¼ A,������ϱ��� B" & _
                    " Where A.ID=B.���ID And A.���=1" & _
                    " And B.����=[4] And (" & strSql & ")" & _
                    " And (A.����ʱ�� is Null Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Order by A.����"
            Else
                'D-ICD-10��������
                If zlCommFun.IsCharChinese(strInput) Then
                    strSql = "���� Like [2]" '���뺺��ʱ,ֻƥ������
                Else
                    strSql = "���� Like [1] Or ���� Like [2] Or " & IIf(gint���� = 0, "����", "�����") & " Like [2]"
                End If
                strSql = _
                    " Select ID,ID as ��ĿID,����,����,����," & IIf(gint���� = 0, "����", "����� as ����") & ",˵��" & _
                    " From ��������Ŀ¼ Where ��� In('D','B') And (" & strSql & ")" & _
                    IIf(mstrSex <> "", " And (�Ա�����=[3] Or �Ա����� is NULL)", "") & _
                    " And (����ʱ�� is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Order by ����"
            End If
        End If
        vRect = zlControl.GetControlRect(txtDiagnose.Hwnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, IIf(optSel(1).Value, "��ϱ���", "��������"), _
            False, "", "", False, False, True, vRect.Left, vRect.Top, txtDiagnose.Height, blnCancel, False, True, _
            strInput & "%", gstrLike & strInput & "%", mstrSex, gint���� + 1, strType)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "û���ҵ�������ƥ������ݡ�", vbInformation, gstrSysName
            End If
            Call zlControl.TxtSelAll(txtDiagnose)
            Exit Sub
        Else
            txtDiagnose.Text = "[" & rsTmp!���� & "]" & Nvl(rsTmp!����)
            txtDiagnose.Tag = Val(rsTmp!��ĿID)
            Set mrsPath = GetPathTable(IIf(mintDiag = E_IX_������, Val(rsTmp!��ĿID), 0), IIf(mintDiag = E_IX_�����, Val(rsTmp!��ĿID), 0), mPati.����ID, 0)
            Call LoadPath(mrsPath)
        End If
    End If
End Sub

Private Sub cmd_Click()
    Dim rsTmp As ADODB.Recordset
    If mblnICD11 Then
        Set rsTmp = ShowILLSelect(Me, "E", mPati.����ID, mstrSex, True, True, , , , 1, True, , 1)
    Else
        If mintDiag = E_IX_����� Then
            '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
            Set rsTmp = ShowILLSelect(Me, "1", mPati.����ID, mstrSex, False, False)
        Else
            'D-ICD-10��������
            Set rsTmp = ShowILLSelect(Me, "D,B", mPati.����ID, mstrSex, False, True)
        End If
    End If
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            txtDiagnose.Text = "[" & rsTmp!���� & "]" & Nvl(rsTmp!����)
            txtDiagnose.Tag = Val(rsTmp!��ĿID)
            Set mrsPath = GetPathTable(IIf(mintDiag = E_IX_������, Val(rsTmp!��ĿID), 0), IIf(mintDiag = E_IX_�����, Val(rsTmp!��ĿID), 0), mPati.����ID, 0)
            Call LoadPath(mrsPath)
        End If
    End If
End Sub

Private Function SaveData() As Boolean
    Dim arrSQL As Variant
    Dim lngסԺ���� As Long, lng��׼סԺ�� As Long
    Dim rsTmp As ADODB.Recordset, rsCriterion As ADODB.Recordset
    Dim strδ������� As String, strδ�������� As String
    Dim lngB As Long, lngE As Long, strUnit As String, strTmp As String, DatCur As Date, lngValue As Long

    Dim bln����ж� As Boolean
    Dim blnTrans As Boolean
    
    Dim dt��Ժʱ�� As Date
    Dim dtDate As Date
    Dim bytDiagSorce As Long
    Dim bytDiagType As Long
    Dim lngPatiDiagID As Long
    Dim strDiagInfo As String   '�������
    
    Dim i As Long
    
    If vsPath.Row < 1 Then
         MsgBox "��ѡ��һ�������ڸò��˵��ٴ�·����", vbInformation + vbOKOnly, gstrSysName
         Exit Function
    End If
    If vsDisease.Row < 1 Then
        MsgBox "��ѡ��һ�������ڸò��˵���ϡ�", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    Else
        If vsDisease.TextMatrix(vsDisease.Row, vsDisease.ColIndex("����")) = "" Then
            MsgBox "���ٴ�·����" & vsPath.TextMatrix(vsPath.Row, vsPath.ColIndex("����")) & "��û����Ч�������Ϣ��", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    If vsDisease.TextMatrix(vsDisease.Row, vsDisease.ColIndex("���")) = "E" And Not mblnICD11 Then
        If mblnHave Then
            MsgBox "�����Ѿ�¼���ICD-11����ϣ���ѡ���ICD-11����ϵ���·����", vbInformation + vbOKOnly, gstrSysName
        Else
            MsgBox "ϵͳ����δ����ICD-11ģʽ����ѡ���ICD-11����ϵ���·����", vbInformation + vbOKOnly, gstrSysName
        End If
        Exit Function
    ElseIf vsDisease.TextMatrix(vsDisease.Row, vsDisease.ColIndex("���")) <> "E" And mblnICD11 Then
        If mblnHave Then
            MsgBox "�����Ѿ�¼��ICD-11����ϣ���ѡ��ICD-11����ϵ���·����", vbInformation + vbOKOnly, gstrSysName
        Else
            MsgBox "ϵͳ��������ICD-11ģʽ����ѡ��ICD-11����ϵ���·����", vbInformation + vbOKOnly, gstrSysName
        End If
        Exit Function
    End If
    mrsPath.Filter = "ID = " & vsPath.RowData(vsPath.Row)
    mPP.·��ID = mrsPath!ID
    mPP.�汾�� = mrsPath!���°汾
    
    With vsDisease
        If Val(.RowData(.Row)) > 0 Then  '����ID
            mrsDisease.Filter = "·��ID=" & mrsPath!ID & " And ����ID =" & .RowData(.Row)
        Else
            mrsDisease.Filter = "·��ID=" & mrsPath!ID & " And ���ID =" & Val(.Cell(flexcpData, .Row, .ColIndex("����")))
        End If
    End With
    
    'mbytDiagSorce=�����Դ1-������2-��Ժ�Ǽǣ�3-��ҳ����;4-����
    bytDiagSorce = 3
    'mbytDiagType=�������:1-��ҽ�������;2-��ҽ��Ժ���;11-��ҽ�������;12-��ҽ��Ժ���
    If InStr(",B,2,", "," & mrsDisease!��� & ",") > 0 Then
        bytDiagType = 12
    Else
        bytDiagType = 2
    End If
    
    
    '���˲������ͣ�����·����Ҳ������Ҫ��
    If Not IsNull(mrsPati!��������) Then
        If mrsPath!�������� <> "��" And mrsPati!�������� <> mrsPath!�������� Then
            MsgBox "��·��Ҫ��Ĳ�������[" & mrsPath!�������� & "]���ʺ��ڸò��˵Ĳ�������[" & mrsPati!�������� & "]", vbInformation, gstrSysName
            strδ�������� = "�������Ͳ�����"
            GoTo UnImport
        End If
    End If
    
    If Not IsNull(mrsPati!��ǰ����) Then
        If mrsPath!���ò��� <> "ͨ��" And mrsPath!���ò��� <> mrsPati!��ǰ���� Then
            MsgBox "��·��[" & mrsPath!���ò��� & "]���ʺ��ڸò��˲���[" & mrsPati!��ǰ���� & "]", vbInformation, gstrSysName
            strδ�������� = "���鲻����"
            GoTo UnImport
        End If
    End If
    If Val(mrsPath!�����Ա�) <> 0 Then
        If Val(mrsPath!�����Ա�) <> IIf(mrsPati!�Ա� = "��", 1, IIf(mrsPati!�Ա� = "Ů", 2, 0)) Then
            MsgBox "��·�����ʺ��ڸò����Ա�[" & mrsPati!�Ա� & "]", vbInformation, gstrSysName
            strδ�������� = "�Ա��ʺ�"
            GoTo UnImport
        End If
    End If
    
    If Not IsNull(mrsPath!��������) And Not IsNull(mrsPati!����) Then
        lngValue = 0
        lngB = Split(mrsPath!��������, "-")(0)
        strTmp = Split(mrsPath!��������, "-")(1)
        lngE = Mid(strTmp, 1, Len(strTmp) - 1)
        strUnit = Mid(strTmp, Len(strTmp))
    
        strTmp = mrsPati!����           '���⣺2��3�µ�
        If strUnit = Mid(strTmp, Len(strTmp)) And IsNumeric(Mid(strTmp, 1, Len(strTmp) - 1)) Or IsNumeric(strTmp) Then
            '��ͬ���䵥λ�����Ƚ�
            lngValue = Val(strTmp)
        ElseIf mcolPati("_pati_birthdate") <> "" Then
            DatCur = zlDatabase.Currentdate
            lngValue = DateDiff(IIf(strUnit = "��", "yyyy", IIf(strUnit = "��", "m", "d")), CDate(mcolPati("_pati_birthdate")), DatCur)
            If lngValue = 0 Then lngValue = 1
        End If
        If lngValue <> 0 Then
            If lngValue < lngB Or lngValue > lngE Then
                MsgBox "��·�����ʺ��ڸò�������[" & mrsPati!���� & "]", vbInformation, gstrSysName
    
                strδ�������� = "���䲻�ʺ�"
                GoTo UnImport
            End If
        End If
    End If
    'סԺ�ղ��ܴ���·���ı�׼סԺ�պ�ȷ������(���û��������ȷ��������������)
    dt��Ժʱ�� = GetPatiInDate(mPati, lngסԺ����)
    dtDate = zlDatabase.Currentdate
    
    If InStr(mrsPath!��׼סԺ��, "-") > 0 Then
        lng��׼סԺ�� = Split(mrsPath!��׼סԺ��, "-")(1)
    Else
        lng��׼סԺ�� = Val(mrsPath!��׼סԺ��)
    End If
    'סԺ��������ȷ��������ֹ����·��;ȷ������δ���û�Ϊ0ʱ,��סԺ�������ڱ�׼סԺ��ʱ��ֹ����·��
    
    If Not CheckPathSend(mPati.����ID, mPati.��ҳID) Then
        If mrsPath!ȷ������ <> 0 Then
            If dtDate > Format(DateAdd("d", Val(mrsPath!ȷ������), dt��Ժʱ��), "yyyy-MM-DD HH:mm:ss") Then
                MsgBox "�ò�������Ժ" & lngסԺ���� & "�죬�����˹涨��ȷ������(" & mrsPath!ȷ������ & "��)��", vbInformation, gstrSysName
                strδ�������� = "����ȷ������"
                GoTo UnImport
            End If
        Else
            If lngסԺ���� > lng��׼סԺ�� Then
                MsgBox "�ò�������Ժ" & lngסԺ���� & "�죬�����˸�·���ı�׼סԺ��(" & lng��׼סԺ�� & "��)��", vbInformation, gstrSysName
                strδ�������� = "������׼סԺ��"
                GoTo UnImport
            End If
        End If
    End If
     
     
    Me.Hide
    bln����ж� = True
    '�ٴ�·������ǰ������ҿ�
    If CreatePlugInOK(P�ٴ�·��Ӧ��) Then
        On Error Resume Next
        bln����ж� = gobjPlugIn.PathImportBefore(glngSys, P�ٴ�·��Ӧ��, mPati.����ID, mPati.��ҳID, mPP.·��ID, mPP.�汾��, bytDiagType, bytDiagSorce, _
        Val(mrsDisease!����id & ""), Val(mrsDisease!���id & ""))
        '����ӿڲ����ڣ���Ӱ��ԭ���߼�
        If Not bln����ж� And Err.Number <> 0 Then bln����ж� = True
        Call zlPlugInErrH(Err, "PathImportBefore")
        Err.Clear: On Error GoTo 0
        If Not bln����ж� Then
            mbln��� = True
            mblnOK = True
            Unload Me
            Exit Function
        End If
    End If
    '
    arrSQL = Array()
    lngPatiDiagID = zlDatabase.GetNextId("������ϼ�¼")
    strDiagInfo = IIf(Val(mrsDisease!����id) > 0, "(" & mrsDisease!������ & ")" & mrsDisease!������, "") & IIf(Val(mrsDisease!���id) > 0, "(" & mrsDisease!����� & ")" & mrsDisease!�����, "")
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_INSERT(" & mPati.����ID & "," & mPati.��ҳID & ",3,NULL," & bytDiagType & "," & _
                        ZVal(mrsDisease!����id) & "," & ZVal(mrsDisease!���id) & ",NULL,'" & _
                        strDiagInfo & "','',0,0," & zlStr.To_Date(Format(dtDate, "yyyy-MM-DD HH:mm:ss"), "ymdhms") & ",'',1,'','',NULL,Null," & lngPatiDiagID & ",NULL,'',''" & _
                        IIf(mrsDisease!��� & "" = "E", ",'E',1,'01'", ",NULL,NULL,NULL") & ",NULL,NULL,NULL,1)"
                        
 
    mblnOK = frmEvaluate.ShowMe(mfrmParent, 0, 1, mPati, mPP, mrsPath!����, bytDiagType, bytDiagSorce, Val(mrsDisease!����id & ""), Val(mrsDisease!���id & ""), 0, , , , arrSQL)
     
    '�ٴ�·������ǰ������ҿ�
    If CreatePlugInOK(P�ٴ�·��Ӧ��) Then
        On Error Resume Next
        Call gobjPlugIn.PathImportAfter(glngSys, P�ٴ�·��Ӧ��, mPati.����ID, mPati.��ҳID, mPP.·��ID, mPP.�汾��)
        Call zlPlugInErrH(Err, "PathImportAfter")
        Err.Clear: On Error GoTo 0
    End If
    
    Unload Me
    Exit Function

UnImport:
    '��Ҫ��ϲű���δ����ԭ��
    Set rsTmp = GetUnImportReson
    rsTmp.Filter = "����='" & strδ�������� & "'"
    If rsTmp.RecordCount = 0 Then
        strδ������� = ""
    Else
        strδ������� = rsTmp!����
    End If
    
    Call SaveUnImport(mPati, mPP, strδ�������, strδ��������, bytDiagType, bytDiagSorce, Val(mrsDisease!����id & ""), Val(mrsDisease!���id & ""))
    mblnOK = True
    Unload Me
End Function


Private Sub vsDisease_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsDisease
        If Me.Visible And NewRow > 0 Then
            If OldRow > 0 Then
                 Set .Cell(flexcpPicture, OldRow, .ColIndex("ѡ��")) = Nothing
            End If
            If .TextMatrix(NewRow, .ColIndex("����")) <> "" Then
                Set .Cell(flexcpPicture, NewRow, .ColIndex("ѡ��")) = imgSrc.ListImages("chkRedUnSquare").Picture
            End If
        End If
    End With
End Sub

Private Sub vsPath_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsPath
        If Me.Visible And NewRow > 0 Then
            If OldRow > 0 Then
                Set .Cell(flexcpPicture, OldRow, .ColIndex("ѡ��")) = Nothing
            End If
            If .TextMatrix(NewRow, .ColIndex("����")) <> "" Then
                Set .Cell(flexcpPicture, NewRow, .ColIndex("ѡ��")) = imgSrc.ListImages("chkRedUnSquare").Picture
            End If
        End If
    End With
End Sub

Private Sub vsPath_Click()
    With vsPath
        If .RowData(.Row) <> "" Then
            Call LoadDisease(.RowData(.Row))
        End If
    End With
End Sub

Private Function IsICDElevent() As Boolean
' �¿�����·��
'    ͨ������ID\��ҳID���Ҳ�����ϼ�¼
'    ��¼Ϊ��
'        ����ICD-11, ��ICD-11
'        δ����ICD-11, ��ICD-10
'    ��¼��Ϊ��
'        ����ICD-10, ��ICD-10
'        ����ICD-11, ��ICD-11
    Dim rsTmp As ADODB.Recordset
    Dim blnResult As Boolean
    
    Set rsTmp = GetDiagType(mPati.����ID, mPati.��ҳID)
    If rsTmp.EOF Then
        blnResult = Mid(gstrICDEleven, 2, 1) = "1"
        mblnHave = False
    Else
        blnResult = (rsTmp!������� & "" = "E")
        mblnHave = True
    End If
    IsICDElevent = blnResult
End Function

