VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmICDElevenEdit 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picICDEleven 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   0
      ScaleHeight     =   5265
      ScaleWidth      =   10035
      TabIndex        =   0
      Top             =   0
      Width           =   10065
      Begin VB.CommandButton cmdAddExPand 
         Appearance      =   0  'Flat
         Caption         =   "������չ��"
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   4800
         Width           =   1100
      End
      Begin VB.CommandButton cmdAddMain 
         Appearance      =   0  'Flat
         Caption         =   "����������"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   4800
         Width           =   1100
      End
      Begin VB.PictureBox picInfectInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3495
         Left            =   0
         ScaleHeight     =   3465
         ScaleWidth      =   10035
         TabIndex        =   2
         Top             =   1200
         Visible         =   0   'False
         Width           =   10065
         Begin VB.ComboBox cboRelation 
            Height          =   300
            ItemData        =   "frmICDElevenEdit.frx":0000
            Left            =   1680
            List            =   "frmICDElevenEdit.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   120
            Width           =   8180
         End
         Begin VB.ListBox lstInfectParts 
            Appearance      =   0  'Flat
            Height          =   2340
            ItemData        =   "frmICDElevenEdit.frx":0004
            Left            =   240
            List            =   "frmICDElevenEdit.frx":000B
            Style           =   1  'Checkbox
            TabIndex        =   3
            Top             =   840
            Width           =   9615
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��Ⱦ�������Ĺ�ϵ"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   10
            Left            =   120
            TabIndex        =   6
            Top             =   180
            Width           =   1440
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ⱦ��λ"
            Height          =   180
            Index           =   128
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   720
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsDiagICDEleven 
         Height          =   1080
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   10065
         _cx             =   17754
         _cy             =   1905
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
         BackColorSel    =   4210752
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmICDElevenEdit.frx":001F
         ScrollTrack     =   -1  'True
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
      Begin VB.Image imgEixt 
         Height          =   360
         Left            =   9480
         Picture         =   "frmICDElevenEdit.frx":0114
         Top             =   4800
         Width           =   360
      End
      Begin VB.Image imgDelete 
         Height          =   360
         Left            =   8160
         Picture         =   "frmICDElevenEdit.frx":07FE
         Top             =   4800
         Width           =   360
      End
      Begin VB.Image imgDown 
         Height          =   360
         Left            =   7440
         Picture         =   "frmICDElevenEdit.frx":0EE8
         Top             =   4800
         Width           =   360
      End
      Begin VB.Image imgUp 
         Height          =   360
         Left            =   6840
         Picture         =   "frmICDElevenEdit.frx":15D2
         Top             =   4800
         Width           =   360
      End
      Begin VB.Image imgSave 
         Height          =   360
         Left            =   8880
         Picture         =   "frmICDElevenEdit.frx":1CBC
         Top             =   4800
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmICDElevenEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmParent As Object
Private mlngModel As Integer, mintDiagType As Integer
Private mRsDiag As ADODB.Recordset
Private mlngX As Long, mlngY As Long, mlngH As Long
Private mstrRelation As String
Private mlng����ID As Long
Private mstr�Ա� As String
Private mblnEnter As Boolean
Private mint���� As Integer
Private mstrLike As String
Private mlng����ID As Long
Private mlng����ID As Long
Private mlngPatiType As Long
Private mDiagTag As String
Private mlng��� As Long
Private mblnSave As Boolean
Private mblnShow As Boolean

Public Function ShowMe(ByRef frmParent As Object, ByVal lngModel As Long, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal lngPatiType As Long, ByVal lng����ID As Long, ByVal str�Ա� As String, ByVal intDiagType As Integer, ByRef rsDiag As ADODB.Recordset, ByVal X As Long, _
    ByVal Y As Long, ByVal txtH As Long, Optional ByRef strRelation As String) As Boolean
'���ܣ�ICD-11��ϲ���������ʾ
'������lngModel-ģ��� lng����ID-����ID lng����ID-(���ﲡ�ˡ��Һ�ID����סԺ���ˡ���ҳID��) lngPatiType-�������ͣ�1-���ﲡ�� 2-סԺ���ˣ�
'      lng����ID-��Ժ����ID str�Ա�-�Ա� intDiagType-������ͣ�1-������� 2-��Ժ��� 3-��Ժ��� 5-Ժ�ڸ�Ⱦ 6-������� 7-�����ж� 11-��ҽ������� 12-��ҽ��Ժ��� 13-��ҽ��Ժ��ϣ�
'      rsDiag-intDiagType��Ӧ����ϼ�¼�� strRelation-��Ⱦ��������ϵ����Ⱦ��λ����ƴ�ӵ��ַ�������ʽΪxx|xx|xx&a��,��strRelation="-1"��ʾ����ʾ��Ⱦ��������ϵ����Ⱦ��λ
    Set mfrmParent = frmParent
    mlngModel = lngModel
    mintDiagType = intDiagType
    Set mRsDiag = rsDiag
    mlngY = Y
    mlngX = X
    mlngH = txtH
    mstrRelation = strRelation
    mlng����ID = lng����ID
    mstr�Ա� = str�Ա�
    mlng����ID = lng����ID
    mlng����ID = lng����ID
    mlngPatiType = lngPatiType
    Me.Show 1, mfrmParent
    strRelation = mstrRelation
    Set rsDiag = mRsDiag
    If mblnSave Then
        ShowMe = True
    Else
        ShowMe = False
    End If
    mblnSave = False
    mblnShow = False
End Function

Private Sub cmdAddExPand_Click()
'���ܣ�������չ��
    Dim LngRow As Long
    Dim i As Long, j As Long
    Dim blnAdd As Boolean
    Dim k As Long
    
    With vsDiagICDEleven
        blnAdd = True
        vsDiagICDEleven.SetFocus
        If .TextMatrix(.Row, Eleven_�������) = "������" Then
            k = 0
            j = 0
        Else
            k = 1
            j = 1
            If .TextMatrix(.Row, Eleven_�������) = "" Then blnAdd = False: Exit Sub
        End If
        '����������¶�Ӧ����չ�����п��е�����¾Ͳ���������������
        For i = .Row To .Rows - 1
            If i + 1 <= .Rows - 1 Then
                If .TextMatrix(i + 1, Eleven_�������) = "��չ��" Or .TextMatrix(i + 1, Eleven_�������) = "֤  ��" Then
                    j = j + 1
                Else
                    Exit For
                End If
                If .TextMatrix(i + 1, Eleven_�������) = "" Then blnAdd = False: Exit For
            End If
        Next
        If Not blnAdd Then Exit Sub
        If k <> 0 Then
            For i = .Row To .FixedRows Step -1
                If i - 1 >= .FixedRows Then
                    If .TextMatrix(i - 1, Eleven_�������) = "��չ��" Or .TextMatrix(i - 1, Eleven_�������) = "֤  ��" Then
                        j = j + 1
                    Else
                        Exit For
                    End If
                    If .TextMatrix(i - 1, Eleven_�������) = "" Then blnAdd = False: Exit For
                End If
            Next
        End If
        If Not blnAdd Then Exit Sub
        '�������Ӧ����չ�벻�ܳ���9������
        If j < 99 Then
            LngRow = .Row + 1: .AddItem "", LngRow
            .TextMatrix(LngRow, Eleven_�������) = IIf(mDiagTag = "��ҽ", "��չ��", "֤  ��")
            .Cell(flexcpBackColor, LngRow, .FixedRows, LngRow, .Cols - 1) = GRD_UNEDITCELL_COLOR
            .Row = LngRow: .Col = Eleven_�������
            vsDiagICDEleven.SetFocus
        Else
            MsgBox "�������Ӧ����չ����ϲ��ܳ���9����", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    Call ChangeVSHeight
End Sub

Private Sub cmdAddMain_Click()
'��������������
    Dim i As Long, j As Long
    Dim LngRow
    Dim blnAdd As Boolean
    
    With vsDiagICDEleven
        blnAdd = True
        vsDiagICDEleven.SetFocus
        '�������б��д����������Ϊ�յ�����������������
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, Eleven_�������) = "������" Then
                j = j + 1
                If .TextMatrix(i, Eleven_�������) = "" Then blnAdd = False: Exit For
            End If
        Next
        If Not blnAdd Then Exit Sub
        'ͬ������ϵ������벻�ܳ���9���
        If j < 99 Then
            LngRow = .Rows: .AddItem "", LngRow
            .TextMatrix(LngRow, Eleven_�������) = "������"
            .Cell(flexcpBackColor, LngRow, .FixedRows, LngRow, .Cols - 1) = &HC0FFC0
            Call ChangeVSHeight
            LngRow = .Rows: .AddItem "", LngRow
            .TextMatrix(LngRow, Eleven_�������) = IIf(mDiagTag = "��ҽ", "��չ��", "֤  ��")
            .Cell(flexcpBackColor, LngRow, .FixedRows, LngRow, .Cols - 1) = GRD_UNEDITCELL_COLOR
            Call ChangeVSHeight
            vsDiagICDEleven.SetFocus
            .Row = LngRow - 1: .Col = Eleven_�������
            .ShowCell .Row, .Col
        Else
            MsgBox "��������ϲ��ܳ���9����", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    Call ChangeVSHeight
End Sub

Private Sub Form_Activate()
    Call ChangeVSHeight
    vsDiagICDEleven.Row = vsDiagICDEleven.FixedRows
    vsDiagICDEleven.Col = Eleven_�������
    vsDiagICDEleven.ShowCell vsDiagICDEleven.Row, vsDiagICDEleven.Col
    vsDiagICDEleven.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim lngScrH As Long
    Me.KeyPreview = True
    If mRsDiag Is Nothing Then Call InitRsICDEleven(mRsDiag)
    If mRsDiag.Fields.Count = 0 Then Call InitRsICDEleven(mRsDiag)
    'ֻ��סԺ��ҳ�Ͳ�����ҳ�Ż���ʾ��Ⱦ��������ϵ�͸�Ⱦ��λ
    If (mlngModel = pסԺҽ��վ Or mlngModel = p��������) And mintDiagType = 5 And mRsDiag.RecordCount >= 1 And mstrRelation <> "-1" Then
        mblnShow = True
    Else
        mblnShow = False
    End If
    If mstrRelation = "-1" Then
        mstrRelation = ""
    End If
    picICDEleven.Height = IIf(mblnShow = False, picICDEleven.Height - picInfectInfo.Height - 100, picICDEleven.Height)
    Me.Height = picICDEleven.Height
    Me.Left = mlngX
    lngScrH = GetSystemMetrics(SM_CYFULLSCREEN) * 15 '��Ļ���ø߶�
    If mlngY + mlngH + Me.Height > lngScrH Then
        Me.Top = mlngY - Me.Height
    Else
        Me.Top = mlngY + mlngH
    End If
    '��ʼ�����
    Call InitTable(vsDiagICDEleven)
    '��ʼ����������
    Call InitData
    '�������
    Call LoadData
End Sub

Private Sub InitTable(ByRef vsTmp As VSFlexGrid)
'���ܣ���ʼ�����
    Dim strHeader As String, strRows As String
    Dim LngRow As Long
    
    With vsTmp
        strHeader = "������,1095,4;��ϱ���,1905,4;�������,6610,1;=����ID;ҽ��IDs;�Ƿ���"
        strRows = Eleven_������� & ",������;" & Eleven_������� & ",��չ��"
        Call Grid.Init(vsTmp, strHeader, strRows, 1, 1)
        .Rows = 3
        .Cell(flexcpBackColor, .FixedRows, Eleven_��ϱ���, .Rows - 1, Eleven_�������) = GRD_UNEDITCELL_COLOR     '����ɫ
        LngRow = .FindRow("������", , Eleven_�������, True)
        .Cell(flexcpBackColor, LngRow, .FixedRows, LngRow, .Cols - 1) = &HC0FFC0
        LngRow = .FindRow("��չ��", , Eleven_�������, True)
        .Cell(flexcpBackColor, LngRow, .FixedRows, LngRow, .Cols - 1) = GRD_UNEDITCELL_COLOR
        .Row = .FixedRows: .Col = Eleven_�������
    End With
End Sub

Private Sub Form_Resize()
    picInfectInfo.Visible = mblnShow
End Sub

Private Sub imgDelete_Click()
'���ܣ�ɾ����ǰ��
    Dim LngRow As Long
    Dim strMsg As String
    Dim lngҽ��ID As Long
    Dim rsTmp As ADODB.Recordset
    Dim str���� As String
    Dim i As Long, j As Long
    
    With vsDiagICDEleven
        LngRow = .Row
        If LngRow + 1 <= .Rows - 1 And .Rows > 3 Then
            If .TextMatrix(LngRow, Eleven_�������) = "������" Then
            '�����ǰ������������ɾ�������뼰��Ӧ����չ��
                For i = LngRow To .Rows - 1
                    If i <= .Rows - 1 Then
                        If .Rows > 3 Then
                            str���� = .TextMatrix(i, Eleven_�������)
                            .Cell(flexcpText, i, .FixedCols, i, .Cols - 1) = ""
                            .Cell(flexcpData, i, .FixedCols, i, .Cols - 1) = Empty
                            .TextMatrix(i, Eleven_�������) = str����
                            If i + 1 <= .Rows - 1 Then
                                For j = Eleven_������� To .Cols - 1
                                    .TextMatrix(i, j) = .TextMatrix(i + 1, j)
                                    .Cell(flexcpData, i, j) = .Cell(flexcpData, i + 1, j)
                                Next
                                .Cell(flexcpBackColor, i, .FixedRows, i, .Cols - 1) = .Cell(flexcpBackColor, i + 1, .FixedRows, i + 1, .Cols - 1)
                                .RowData(i) = .RowData(i + 1)
                                If .TextMatrix(i + 1, Eleven_�������) = "������" Then
                                    .RemoveItem i + 1
                                    Exit For
                                Else
                                    .RemoveItem i + 1
                                    i = i - 1
                                End If
                            Else
                                .RemoveItem i
                            End If
                        Else
                            str���� = .TextMatrix(i, Eleven_�������)
                            .Cell(flexcpText, i, .FixedCols, i, .Cols - 1) = ""
                            .Cell(flexcpData, i, .FixedCols, i, .Cols - 1) = Empty
                            .TextMatrix(i, Eleven_�������) = str����
                            .Cell(flexcpBackColor, i, .FixedCols, i, .Cols - 1) = IIf(i = 1, &HC0FFC0, GRD_UNEDITCELL_COLOR)
                        End If
                    End If
                Next
            Else
                '�����ǰ������չ����ɾ����ǰ��
                str���� = .TextMatrix(LngRow, Eleven_�������)
                .Cell(flexcpText, LngRow, .FixedCols, LngRow, .Cols - 1) = ""
                .Cell(flexcpData, LngRow, .FixedCols, LngRow, .Cols - 1) = Empty
                .TextMatrix(LngRow, Eleven_�������) = str����
                If .TextMatrix(LngRow - 1, Eleven_�������) = "������" Then
                    If LngRow + 1 <= .Rows - 1 Then
                        If .TextMatrix(LngRow + 1, Eleven_�������) = "��չ��" Or .TextMatrix(LngRow + 1, Eleven_�������) = "֤  ��" Then
                            .RemoveItem LngRow
                        End If
                    End If
                Else
                    For j = Eleven_������� To .Cols - 1
                        .TextMatrix(LngRow, j) = .TextMatrix(LngRow + 1, j)
                        .Cell(flexcpData, i, j) = .Cell(flexcpData, LngRow + 1, j)
                    Next
                    .Cell(flexcpBackColor, LngRow, .FixedRows, LngRow, .Cols - 1) = .Cell(flexcpBackColor, LngRow + 1, .FixedRows, LngRow + 1, .Cols - 1)
                    .RowData(LngRow) = .RowData(LngRow + 1)
                    .RemoveItem LngRow + 1
                End If
            End If
        Else
            If .Rows > 3 Then
                For i = LngRow To .Rows - 1
                    If i <= .Rows - 1 Then
                        If i - 1 <= .Rows - 1 Then
                            If .TextMatrix(i - 1, Eleven_�������) = "������" Then
                                 str���� = .TextMatrix(i, Eleven_�������)
                                .Cell(flexcpText, i, .FixedCols, i, .Cols - 1) = ""
                                .Cell(flexcpData, i, .FixedCols, i, .Cols - 1) = Empty
                                .TextMatrix(i, Eleven_�������) = str����
                                 .Cell(flexcpBackColor, i, .FixedCols, i, .Cols - 1) = IIf(i = 1, &HC0FFC0, GRD_UNEDITCELL_COLOR)
                            Else
                                .RemoveItem i
                                i = i - 1
                            End If
                        Else
                            .RemoveItem i
                            i = i - 1
                        End If
                    End If
                Next
            Else
                For i = LngRow To .Rows - 1
                    str���� = .TextMatrix(i, Eleven_�������)
                    .Cell(flexcpText, i, .FixedCols, i, .Cols - 1) = ""
                    .Cell(flexcpData, i, .FixedCols, i, .Cols - 1) = Empty
                    .TextMatrix(i, Eleven_�������) = str����
                    .Cell(flexcpBackColor, i, .FixedCols, i, .Cols - 1) = IIf(i = 1, &HC0FFC0, GRD_UNEDITCELL_COLOR)
                Next
            End If
        End If
    End With
    Call ChangeVSHeight
End Sub

Private Sub imgDown_Click()
'���ܣ���ǰ��������
    With vsDiagICDEleven
        '�����ǰ������չ������һ��Ϊ�����������������ƶ�
        If .Row + 1 <= .Rows - 1 And (.TextMatrix(.Row, Eleven_�������) = "��չ��" Or .TextMatrix(.Row, Eleven_�������) = "֤  ��") Then
            If .TextMatrix(.Row + 1, Eleven_�������) = "������" Then Exit Sub
        End If
        '�����ƶ�
        Call MoveCurrRow(.TextMatrix(.Row, Eleven_�������), .Row, -1)
    End With
End Sub

Private Sub imgEixt_Click()
    Unload Me
End Sub

Private Sub imgSave_Click()
'�������¼�������
    If Not CheckDate Then Exit Sub
    If Not SaveData Then
        Exit Sub
    Else
        Unload Me
    End If
End Sub

Private Function CheckDate() As Boolean
'�������б�¼�������
    Dim i As Long, j As Long
    Dim blnHaveDaig As Boolean
    Dim str������� As String
    Dim lngColor As Long
    
    With vsDiagICDEleven
        '����Ƿ����������ͬ�����������
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, Eleven_�������)) <> "" And .TextMatrix(i, Eleven_�������) = "������" Then
                If i <> .Rows - 1 Then
                     For j = i + 1 To .Rows - 1
                        If .TextMatrix(j, Eleven_�������) = "������" Then
                            If Trim(.TextMatrix(i, Eleven_�������)) = Trim(.TextMatrix(j, Eleven_�������)) Then
                                .Row = i: .Col = Eleven_�������
                                lngColor = .CellBackColor: .CellBackColor = &HC0C0FF
                                Call .ShowCell(.Row, .Col)
                                MsgBox "���ִ���������ͬ�������Ϣ��", vbInformation, gstrSysName
                                .CellBackColor = lngColor
                                str������� = .TextMatrix(i, Eleven_�������)
                                vsDiagICDEleven.SetFocus: Exit Function
                            End If
                        End If
                     Next
                End If
            End If
        Next
        '����������Ӧ����չ���Ƿ����������ͬ�����
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, Eleven_�������)) <> "" And (.TextMatrix(i, Eleven_�������) = "��չ��" Or .TextMatrix(i, Eleven_�������) = "֤  ��") Then
                If i <> .Rows - 1 Then
                     For j = i + 1 To .Rows - 1
                        If .TextMatrix(j, Eleven_�������) = "������" Then
                            Exit For
                        Else
                            If Trim(.TextMatrix(i, Eleven_�������)) = Trim(.TextMatrix(j, Eleven_�������)) Then
                                .Row = i: .Col = Eleven_�������
                                lngColor = .CellBackColor: .CellBackColor = &HC0C0FF
                                Call .ShowCell(.Row, .Col)
                                MsgBox "���ִ���������ͬ�������Ϣ", vbInformation, gstrSysName
                                .CellBackColor = lngColor
                                str������� = .TextMatrix(i, Eleven_�������)
                                vsDiagICDEleven.SetFocus: Exit Function
                            End If
                        End If
                     Next
                End If
            End If
        Next
        '�����չ���Ƿ��Ӧ��������
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, Eleven_�������) = "������" Then
                If i <> .Rows - 1 Then
                     For j = i + 1 To .Rows - 1
                        If .TextMatrix(j, Eleven_�������) = "������" Then
                            Exit For
                        Else
                            If .TextMatrix(j, Eleven_�������) <> "" Then blnHaveDaig = True
                        End If
                    Next
                    If blnHaveDaig And .TextMatrix(i, Eleven_�������) = "" Then
                        .Row = i: .Col = Eleven_�������
                        lngColor = .CellBackColor: .CellBackColor = &HC0C0FF
                        Call .ShowCell(.Row, .Col)
                        MsgBox "��չ��û�ж�Ӧ�������룬���飡", vbInformation, gstrSysName
                        .CellBackColor = lngColor
                        str������� = .TextMatrix(i, Eleven_�������)
                        vsDiagICDEleven.SetFocus: Exit Function
                    End If
                End If
            End If
        Next
    End With
    CheckDate = True
End Function

Private Function SaveData() As Boolean
    '�������¼�������
    Dim strValues As String
    Dim j As Long, i As Long
    Dim k As Long
    Dim rsDiag As ADODB.Recordset
    
    Call InitRsICDEleven(rsDiag)
    
    With vsDiagICDEleven
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, Eleven_�������) <> "" Then
                If .TextMatrix(i, Eleven_�������) = "������" Then
                    k = 0
                    j = j + 1
                End If
                If .TextMatrix(i, Eleven_�������) = "��չ��" Or .TextMatrix(i, Eleven_�������) = "֤  ��" Then
                    k = k + 1
                End If
                rsDiag.AddNew Array("��Ϣ��", "�������", "��ϱ���", "�������", "����ID", "ҽ��IDs", "Tag", "IndexEx", "���"), Array("ICD-11", .TextMatrix(i, Eleven_�������), .TextMatrix(i, Eleven_��ϱ���), .TextMatrix(i, Eleven_�������), Val(.TextMatrix(i, Eleven_����ID)), .TextMatrix(i, Eleven_ҽ��IDs), mDiagTag, IIf(.TextMatrix(i, Eleven_�������) = "������", IIf(j <= 9, "0" & j, "" & j), IIf(j <= 9, "0" & j, "" & j) & IIf(k <= 9, "0" & k, "" & k)), mlng���)
            End If
        Next
        Set mRsDiag = zlDatabase.CopyNewRec(rsDiag)
    End With
    '��Ⱦ��������ϵ����Ⱦ��λ
    If picInfectInfo.Visible Then
        For j = 0 To lstInfectParts.ListCount - 1
            If lstInfectParts.Selected(j) = True Then
                strValues = strValues & "|" & lstInfectParts.ItemData(j)
            End If
        Next
        If strValues <> "" Then
            strValues = Mid(strValues, 2)
        End If
        strValues = strValues & "&" & zlcommfun.GetNeedName(cboRelation.Text, "-")
    End If
    mstrRelation = strValues
    SaveData = True
    mblnSave = SaveData
End Function

Private Sub imgUp_Click()
'�����ƶ�
    With vsDiagICDEleven
        '�����ǰ������չ������һ��Ϊ���������������ƶ�
        If .Row - 1 >= .FixedRows And (.TextMatrix(.Row, Eleven_�������) = "��չ��" Or .TextMatrix(.Row, Eleven_�������) = "֤  ��") Then
            If .TextMatrix(.Row - 1, Eleven_�������) = "������" Then Exit Sub
        End If
        '�����ƶ�
        Call MoveCurrRow(.TextMatrix(.Row, Eleven_�������), .Row, 1)
    End With
End Sub

Private Sub picICDEleven_Resize()
    If picInfectInfo.Visible Then
        picInfectInfo.Top = vsDiagICDEleven.Top + vsDiagICDEleven.Height + 100
        cmdAddMain.Top = picInfectInfo.Top + picInfectInfo.Height + 100
        cmdAddExPand.Top = picInfectInfo.Top + picInfectInfo.Height + 100
        imgUp.Top = picInfectInfo.Top + picInfectInfo.Height + 100
        imgDown.Top = picInfectInfo.Top + picInfectInfo.Height + 100
        imgDelete.Top = picInfectInfo.Top + picInfectInfo.Height + 100
        imgSave.Top = picInfectInfo.Top + picInfectInfo.Height + 100
        imgEixt.Top = picInfectInfo.Top + picInfectInfo.Height + 100
    Else
        cmdAddMain.Top = vsDiagICDEleven.Top + vsDiagICDEleven.Height + 100
        cmdAddExPand.Top = vsDiagICDEleven.Top + vsDiagICDEleven.Height + 100
        imgUp.Top = vsDiagICDEleven.Top + vsDiagICDEleven.Height + 100
        imgDown.Top = vsDiagICDEleven.Top + vsDiagICDEleven.Height + 100
        imgDelete.Top = vsDiagICDEleven.Top + vsDiagICDEleven.Height + 100
        imgSave.Top = vsDiagICDEleven.Top + vsDiagICDEleven.Height + 100
        imgEixt.Top = vsDiagICDEleven.Top + vsDiagICDEleven.Height + 100
    End If
End Sub

Private Sub LoadData()
'���ܣ���ʾ����
    Dim i As Long, j As Long, n As Long
    Dim lngRows As Long
    Dim strArrRelation As Variant
    Dim strTmp As String
    
    With vsDiagICDEleven
        For i = .FixedRows To .Rows - 1
            If i <= .Rows - 1 Then
                If i <> 1 And i <> 2 Then
                 .RemoveItem i
                End If
            End If
        Next
        For i = .FixedRows To .Rows - 1
            If i = 1 Then
                .TextMatrix(i, Eleven_�������) = "������"
                .Cell(flexcpBackColor, i, .FixedRows, i, .Cols - 1) = &HC0FFC0
            Else
                .TextMatrix(i, Eleven_�������) = "��չ��"
                .Cell(flexcpBackColor, i, .FixedRows, i, .Cols - 1) = GRD_UNEDITCELL_COLOR
            End If
            .TextMatrix(i, Eleven_����ID) = ""
            .TextMatrix(i, Eleven_�������) = ""
            .Cell(flexcpData, i, Eleven_�������) = .TextMatrix(i, Eleven_�������)
            .TextMatrix(i, Eleven_��ϱ���) = ""
            .TextMatrix(i, Eleven_ҽ��IDs) = ""
            .TextMatrix(i, Eleven_�Ƿ���) = ""
            .RowData(i) = ""
        Next
        If Not mRsDiag.EOF Then
            For i = 0 To mRsDiag.RecordCount - 1
                mDiagTag = mRsDiag!Tag
                If .TextMatrix(n, Eleven_�������) = "������" And "" & mRsDiag!������� = "������" Then
                    n = n + 1
                    .AddItem "", n
                    .TextMatrix(n, Eleven_�������) = IIf(mDiagTag = "��ҽ", "��չ��", "֤  ��")
                    .Cell(flexcpBackColor, n, .FixedRows, n, .Cols - 1) = GRD_UNEDITCELL_COLOR
                End If
                n = n + 1
                If n > .Rows - 1 Then
                    .AddItem "", n
                End If
                .TextMatrix(n, Eleven_�������) = "" & mRsDiag!�������
                .TextMatrix(n, Eleven_��ϱ���) = "" & mRsDiag!��ϱ���
                .TextMatrix(n, Eleven_�������) = "" & mRsDiag!�������
                .TextMatrix(n, Eleven_����ID) = "" & mRsDiag!����id
                .TextMatrix(n, Eleven_ҽ��IDs) = "" & mRsDiag!ҽ��IDs
                .TextMatrix(n, Eleven_�Ƿ���) = ""
                If "" & mRsDiag!������� = "������" Then
                    .Cell(flexcpBackColor, n, .FixedRows, n, .Cols - 1) = &HC0FFC0
                Else
                    .Cell(flexcpBackColor, n, .FixedRows, n, .Cols - 1) = GRD_UNEDITCELL_COLOR
                End If
                mlng��� = Val("" & mRsDiag!���)
                mRsDiag.MoveNext
                Call ChangeVSHeight
            Next
        End If
        If mintDiagType = 11 Or mintDiagType = 12 Or mintDiagType = 13 Then
            mDiagTag = "��ҽ"
        Else
            mDiagTag = "��ҽ"
        End If
        For i = .FixedRows To .Rows - 1
            If i + 1 <= .Rows - 1 Then
                If .TextMatrix(i, Eleven_�������) = "������" Then
                    If .TextMatrix(i + 1, Eleven_�������) <> "��չ��" And .TextMatrix(i + 1, Eleven_�������) <> "֤  ��" Then
                        .AddItem "", i + 1
                        .TextMatrix(i + 1, Eleven_�������) = IIf(mDiagTag = "��ҽ", "��չ��", "֤  ��")
                        .Cell(flexcpBackColor, i + 1, .FixedRows, i + 1, .Cols - 1) = GRD_UNEDITCELL_COLOR
                    End If
                End If
            Else
                If .TextMatrix(i, Eleven_�������) = "������" Then
                    .AddItem "", i + 1
                    .TextMatrix(i + 1, Eleven_�������) = IIf(mDiagTag = "��ҽ", "��չ��", "֤  ��")
                    .Cell(flexcpBackColor, i + 1, .FixedRows, i + 1, .Cols - 1) = GRD_UNEDITCELL_COLOR
                End If
            End If
            If .TextMatrix(i, Eleven_�������) = "��չ��" Then
                .TextMatrix(i, Eleven_�������) = IIf(mDiagTag = "��ҽ", "��չ��", "֤  ��")
            End If
        Next
    End With
    
    If mstrRelation <> "" Then
        strTmp = Mid(mstrRelation, 1, InStr(mstrRelation, "&") - 1)
        If strTmp <> "" Then
            With lstInfectParts
                strArrRelation = Split(strTmp, "|")
                For j = 0 To .ListCount - 1
                    For i = LBound(strArrRelation) To UBound(strArrRelation)
                        If .ItemData(j) = strArrRelation(i) Then
                            .Selected(j) = True: Exit For
                        End If
                    Next
                Next
                .ListIndex = -1
            End With
        End If
        strTmp = Mid(mstrRelation, InStr(mstrRelation, "&") + 1)
        If strTmp <> "" Then
            Call Cbo.SeekIndex(cboRelation, strTmp)
        End If
    End If
End Sub

Private Sub InitData()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    mint���� = Val(zlDatabase.GetPara("���뷽ʽ"))
    mstrLike = IIf(zlDatabase.GetPara("����ƥ��") = "0", "%", "")
    
    cboRelation.AddItem " "
    cboRelation.AddItem "0-ֱ��"
    cboRelation.AddItem "1-���"
    cboRelation.AddItem "2-��"
    cboRelation.ListIndex = -1
    
    strSql = "Select RowNum As ID, ����, ����, ����, ȱʡ��־ ȱʡ From ��Ⱦ��λ"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ��Ⱦ��λ")
    With lstInfectParts
        If Not rsTmp.EOF Then
            lstInfectParts.Clear
            rsTmp.Sort = "����,����"
            Do While Not rsTmp.EOF
                .AddItem rsTmp!����
                .ItemData(.NewIndex) = Val(rsTmp!����)
                rsTmp.MoveNext
            Loop
        End If
    End With
End Sub

Private Function ChangeVSHeight() As Boolean
'���ܣ���PictureBox �������VSF�Ĵ�С
    Dim i As Long
    Dim lngOldVSFHeight As Long
    Dim lngRows As Long
    Dim lngVSFHeight As Long
    Dim lngRowHeight As Long
    Dim lngMaxHeight As Long
    Dim lngShowRows As Long
    Dim lngScrH As Long
    Dim lngLastHeight As Long
    
    lngRowHeight = IIf(vsDiagICDEleven.RowHeightMax < vsDiagICDEleven.RowHeightMin, vsDiagICDEleven.RowHeightMin, vsDiagICDEleven.RowHeightMax)
    lngOldVSFHeight = vsDiagICDEleven.Height
    lngRows = vsDiagICDEleven.Rows
    For i = 0 To vsDiagICDEleven.Rows - 1
        lngVSFHeight = lngVSFHeight + vsDiagICDEleven.RowHeight(i)
        lngShowRows = lngShowRows + 1
    Next
    lngVSFHeight = IIf(lngVSFHeight < lngShowRows * lngRowHeight, lngShowRows * lngRowHeight + 30, lngVSFHeight)
    lngScrH = GetSystemMetrics(SM_CYFULLSCREEN) * 15 '��Ļ���ø߶�
    lngLastHeight = picICDEleven.Height + (lngVSFHeight - lngOldVSFHeight)
    If lngLastHeight > mlngY - 50 Then
        If mlngH + lngLastHeight > lngScrH - mlngY - mlngH Then
            vsDiagICDEleven.Height = vsDiagICDEleven.Height
        Else
            vsDiagICDEleven.Height = lngVSFHeight
        End If
    Else
        vsDiagICDEleven.Height = lngVSFHeight
    End If
    If vsDiagICDEleven.Height - lngOldVSFHeight <> 0 Then
        picICDEleven.Height = picICDEleven.Height + (vsDiagICDEleven.Height - lngOldVSFHeight)
    End If
    Me.Height = picICDEleven.Height
    If mlngY + mlngH + Me.Height > lngScrH Then
        Me.Top = mlngY - Me.Height
    Else
        Me.Top = mlngY + mlngH
    End If
End Function

Private Sub vsDiagICDEleven_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long, j As Long, k As Long
    Dim LngRow As Long, LngCol As Long
    Dim blnDel As Boolean
    
    LngRow = Row
    LngCol = Col
    With vsDiagICDEleven
        If LngCol = Eleven_������� Then
            ' .EditText = "" �ų���Ԫ�������ݲ����س���״��
            If (LngCol = Eleven_������� And .TextMatrix(LngRow, Eleven_��ϱ���) <> "" Or LngCol = Eleven_��ϱ��� And .TextMatrix(LngRow, Eleven_�������) <> "") And .EditText = "" And .Cell(flexcpData, LngRow, LngCol) <> "" Then
                If .TextMatrix(LngRow, Eleven_�������) = "������" Then
                    For i = LngRow + 1 To .Rows - 1
                        If .TextMatrix(i, Eleven_�������) = "������" Then
                            Exit For
                        Else
                            If .TextMatrix(i, Eleven_�������) <> "" Then
                                If MsgBox("�Ƿ���ɾ���������ͬʱɾ����Ӧ����չ�룿����ǣ�ͬ��ɾ����Ӧ����չ�룻�������ɾ��Ӧ����չ�롣", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then blnDel = True
                                Exit For
                            End If
                        End If
                    Next
                End If
                If Not blnDel Then
                    .TextMatrix(LngRow, LngCol) = .Cell(flexcpData, LngRow, LngCol)
                    .Cell(flexcpText, LngRow, .FixedCols, LngRow, .Cols - 1) = ""
                    .Cell(flexcpData, LngRow, .FixedCols, LngRow, .Cols - 1) = Empty
                    For j = .FixedCols To .Cols - 1
                        .TextMatrix(LngRow, j) = ""
                    Next
                Else
                    For i = LngRow To .Rows - 1
                        If i <= .Rows - 1 Then
                            k = k + 1
                            If .TextMatrix(i, Eleven_�������) = "������" And i <> LngRow Then Exit For
                            .TextMatrix(i, LngCol) = .Cell(flexcpData, i, LngCol)
                            .Cell(flexcpText, i, .FixedCols, i, .Cols - 1) = ""
                            .Cell(flexcpData, i, .FixedCols, i, .Cols - 1) = Empty
                            For j = .FixedCols To .Cols - 1
                                .TextMatrix(i, j) = ""
                            Next
                            If k > 2 Then
                                .RemoveItem i
                                i = i - 1
                            End If
                        End If
                    Next
                End If
            End If
        End If
        Call ChangeVSHeight
        Call vsDiagICDEleven_AfterRowColChange(-1, -1, .Row, .Col)
    End With
End Sub

Private Sub vsDiagICDEleven_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngNewRow As Long, lngNewCol As Long
    
    lngNewRow = NewRow: lngNewCol = NewCol
    If lngNewRow = -1 Or lngNewCol = -1 Then Exit Sub
    If vsDiagICDEleven.Editable = flexEDNone Then Exit Sub
    
    With vsDiagICDEleven
        If Not ICDElevenEditable(lngNewRow, lngNewCol) Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .ComboList = ""
            .FocusRect = flexFocusSolid
            Select Case lngNewCol
                Case Eleven_�������
                    .ComboList = "..."
                Case Eleven_��ϱ���
                    
                Case Else
                    .ComboList = ""
            End Select
        End If
    End With
End Sub

Private Sub vsDiagICDEleven_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not ICDElevenEditable(Row, Col) Then
        Cancel = True
    End If
End Sub

Private Sub vsDiagICDEleven_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> Eleven_������� Then Cancel = True
End Sub

Private Sub vsDiagICDEleven_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
     Dim i As Long
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim vbPoint As POINTAPI
    Dim blnCancel As Boolean

    Call CreatePublicAdvice
    With vsDiagICDEleven
        Select Case Col
            Case Eleven_�������
                If .TextMatrix(Row, Eleven_�������) = "������" Then
                    If gobjPublicAdvice Is Nothing Then Exit Sub
                    Set rsTmp = gobjPublicAdvice.ShowILLSelect(Me, "E", mlng����ID, , False, False, , , mlngPatiType, 1, True, mintDiagType)
                Else
                    If gobjPublicAdvice Is Nothing Then Exit Sub
                    Set rsTmp = gobjPublicAdvice.ShowILLSelect(Me, "E", mlng����ID, , False, False, , , mlngPatiType, 1, False, mintDiagType)
                End If
                Call SetICDElvenInput(rsTmp, Row, Col)
        End Select
    End With
End Sub

Private Sub SetICDElvenInput(ByVal rsTmp As ADODB.Recordset, ByVal LngRow As Long, ByVal LngCol As Long)
    Dim i As Long
    With vsDiagICDEleven
        If rsTmp Is Nothing Then Exit Sub
        If rsTmp.EOF Then
            .EditText = .TextMatrix(LngRow, Eleven_�������)
        Else
            For i = 0 To rsTmp.RecordCount - 1
                .TextMatrix(LngRow, Eleven_�������) = rsTmp!���� & ""
                .EditText = .TextMatrix(LngRow, Eleven_�������)
                .Cell(flexcpData, LngRow, Eleven_�������) = .TextMatrix(LngRow, Eleven_�������)
                .TextMatrix(LngRow, Eleven_��ϱ���) = rsTmp!���� & ""
                .TextMatrix(LngRow, Eleven_����ID) = rsTmp!����id & ""
                .TextMatrix(LngRow, Eleven_�Ƿ���) = "" & rsTmp!�Ƿ���
                .RowData(LngRow) = rsTmp!��ĿID & ""
                rsTmp.MoveNext
            Next
        End If
    End With
End Sub

Private Sub vsDiagICDEleven_DblClick()
    Call vsDiagICDEleven_KeyPress(vbKeySpace)
End Sub

Private Sub vsDiagICDEleven_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Call imgDelete_Click
    Else
        vsDiagICDEleven_KeyPress KeyCode
    End If
End Sub

Private Sub vsDiagICDEleven_KeyPress(KeyAscii As Integer)
    Dim LngRow As Long, LngCol As Long
    With vsDiagICDEleven
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call EnterCellDiag(.Row, .Col)
        Else
            If Not ICDElevenEditable(.Row, .Col) Then Exit Sub
            Select Case .Col
                Case Eleven_�������
                    If KeyAscii = Asc("*") Then
                        KeyAscii = 0
                        Call vsDiagICDEleven_CellButtonClick(.Row, .Col)
                    Else
                        LngRow = .Row
                        LngCol = .Col
                        .ComboList = ""
                    End If
            End Select
        End If
    End With
End Sub

Private Sub EnterCellDiag(ByVal LngRow As Long, ByVal LngCol As Long)
    Dim i As Long, j As Long

    With vsDiagICDEleven
        '����һ��Ԫ��ʼѭ������
        If LngRow < .FixedRows Then LngRow = .FixedRows
        For i = LngRow To .Rows - 1
            For j = IIf(i = LngRow, LngCol + 1, Eleven_��ϱ���) To Eleven_�������
                If Not .ColHidden(j) Then
                    If ICDElevenEditable(i, j) And .ColWidth(j) <> 0 Then Exit For
                End If
            Next
            If j <= Eleven_������� Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
        ElseIf i = .Rows And j > Eleven_������� And .TextMatrix(.Rows - 1, Eleven_�������) <> "" Then
            If .TextMatrix(.Row, Eleven_�������) = "��չ��" Or .TextMatrix(.Row, Eleven_�������) = "֤  ��" Then
                If .TextMatrix(.Row - 1, Eleven_�������) <> "" Then
                    .Rows = .Rows + 1
                    Call ChangeVSHeight
                    .TextMatrix(.Rows - 1, Eleven_�������) = .TextMatrix(.Rows - 2, Eleven_�������)
                    .Cell(flexcpBackColor, .Rows - 1, .FixedRows, .Rows - 1, .Cols - 1) = .Cell(flexcpBackColor, .Rows - 2, .FixedRows, .Rows - 2, .Cols - 1)
                    .Row = .Rows - 1: .Col = Eleven_�������
                End If
            End If
        Else
            Call zlcommfun.PressKey(vbKeyTab): mblnEnter = True
        End If
    End With
End Sub


Public Function ICDElevenEditable(ByVal LngRow As Long, ByVal LngCol As Long) As Boolean
    Dim blnJudge As Boolean
    Dim i As Long

    With vsDiagICDEleven
        If LngCol <> Eleven_������� Then Exit Function
        Select Case LngCol
            Case Eleven_�������
                If .TextMatrix(LngRow, Eleven_�������) <> "" Then
'                    For i = .FixedRows To .Rows - 1
'                        If .TextMatrix(i, Eleven_�������) = .TextMatrix(lngRow, Eleven_�������) And .TextMatrix(i, Eleven_�������) = "" And .TextMatrix(lngRow, Eleven_�������) = "������" Then
'                            blnJudge = True
'                        End If
'                    Next
'                    If blnJudge Then Exit Function
                Else
                    If LngRow - 1 >= .FixedRows Then
                        If .TextMatrix(LngRow - 1, Eleven_�������) = "������" Then
                            If .TextMatrix(LngRow - 1, Eleven_�������) = "" Then blnJudge = True
                            If blnJudge Then Exit Function
                        End If
                    End If
                End If
        End Select
    End With
    ICDElevenEditable = True
End Function

Private Sub vsDiagICDEleven_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnInputCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    Dim strDiagType As String
    Dim strTag As String
    Dim strTmp As String
    Dim i As Long
    Dim LngRow As Long, LngCol As Long

    With vsDiagICDEleven
        LngRow = Row: LngCol = Col
        Select Case LngCol
            Case Eleven_�������
                strTmp = .TextMatrix(LngRow, Eleven_�������)
                If .TextMatrix(LngRow, Eleven_�������) = "������" Then
                    strDiagType = IIf(mDiagTag = "��ҽ", decode(mintDiagType, "6", "',2,'", "7", "',23,'", "',1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,27,'"), "',26,'")
                Else
                    strDiagType = IIf(mDiagTag = "��ҽ", "',28,'", "',26,'")
                End If
                If .EditText = "" And .TextMatrix(.Row, .Col) <> "" Then
                    .EditText = ""
                ElseIf .EditText = .Cell(flexcpData, LngRow, LngCol) Then
                    If mblnEnter Then Call EnterCellDiag(LngRow, LngCol)
                ElseIf .TextMatrix(LngRow, Eleven_��ϱ���) <> "" And .Cell(flexcpData, LngRow, LngCol) <> "" Then
                    strInput = UCase(.EditText)
                    strSql = GetICDElevenSql(strInput, mstr�Ա�, IIf(.TextMatrix(LngRow, Eleven_�������) = "������", 0, 1), strDiagType)
                    vPoint = GetCoordPos(.hwnd, .Left + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, strTag, _
                        False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                        strInput & "%", mstrLike & strInput & "%", "E", mstr�Ա�, mint���� + 1, strInput, UserInfo.ID, mlng����ID)
                    If blnInputCancel Then
                        Cancel = True
                        .EditText = strTmp
                        .TextMatrix(LngRow, Eleven_�������) = strTmp
                        .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, Eleven_�������)
                    Else
                        If rsTmp Is Nothing Then
                             Cancel = True
                            .EditText = strTmp
                            .TextMatrix(LngRow, Eleven_�������) = strTmp
                            .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, Eleven_�������)
                        Else
                             Call SetICDElvenInput(rsTmp, LngRow, LngCol)
                        End If
                    End If
                Else
                    strInput = UCase(.EditText)
                    strSql = GetICDElevenSql(strInput, mstr�Ա�, IIf(.TextMatrix(LngRow, Eleven_�������) = "������", 0, 1), strDiagType)
                    vPoint = GetCoordPos(.hwnd, .Left + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, strTag, _
                        False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                        strInput & "%", mstrLike & strInput & "%", "E", mstr�Ա�, mint���� + 1, strInput, UserInfo.ID, mlng����ID)
                    If blnInputCancel Then
                        Cancel = True
                        .EditText = strTmp
                        .TextMatrix(LngRow, Eleven_�������) = strTmp
                        .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, Eleven_�������)
                    Else
                        If Not rsTmp Is Nothing Then
                            Call SetICDElvenInput(rsTmp, LngRow, LngCol)
                        Else
                            Cancel = True
                            .EditText = strTmp
                            .TextMatrix(LngRow, Eleven_�������) = strTmp
                            .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, Eleven_�������)
                        End If
                    End If
                End If
        End Select
    End With
End Sub

Private Function GetICDElevenSql(ByVal strInput As String, ByRef str�Ա� As String, ByVal intType As Integer, Optional ByVal strOtherInfo As String) As String
    Dim strSql As String
    Dim lng������� As Long, lng֤����� As Long
    Dim rsTmp As ADODB.Recordset
    
    If strOtherInfo = "',26,'" Then
        strSql = "Select ��� From ����������� Where �½� = '26' And ���� = '��ͳҽѧ������TM1��' And ���� = 'L1-SA0'"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "�����������")
        If Not rsTmp.EOF Then
            lng������� = Val("" & rsTmp!���)
        End If
        
        strSql = "Select ��� From ����������� Where �½� = '26' And ���� = '��ͳҽѧ֤��TM1��' And ���� = 'L1-SE7'"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "�����������")
        If Not rsTmp.EOF Then
            lng֤����� = Val("" & rsTmp!���)
        End If
    End If
    If zlcommfun.IsCharChinese(strInput) Then
        strSql = "A.���� Like [2]" & IIf(intType = 0, IIf(lng������� <> 0 And lng֤����� <> 0, " And C.���>= " & lng������� & " And C.���<" & lng֤�����, ""), IIf(lng֤����� <> 0, " And C.���>" & lng֤�����, "")) '���뺺��ʱֻƥ������
    Else
        strSql = "A.���� Like [1] Or A.���� Like [2] Or " & IIf(mint���� = 0, "A.����", "A.�����") & " Like [2]" & IIf(intType = 0, IIf(lng������� <> 0 And lng֤����� <> 0, " And C.���>= " & lng������� & " And C.���<" & lng֤�����, ""), IIf(lng֤����� <> 0, " And C.���>" & lng֤�����, ""))
    End If
    strSql = "Select A.Id, A.Id ��ĿID,A.����,A.����," & IIf(mint���� = 0, "A.����", "A.�����") & " as ����,  C.�Ƿ���,A.���� ��������, A.Id ����id,A.��� �������" & vbNewLine & _
        "From ��������Ŀ¼ A, ����������� C" & vbNewLine & _
        "Where A.����id = C.Id(+) And A.�½�=C.�½�(+) " & IIf(strOtherInfo <> "", " And Instr(" & strOtherInfo & "," & " ',' || A.�½� || ',')>0", "") & " And Instr([3],A.���)>0 And (" & strSql & ")" & _
        IIf(str�Ա� <> "", " And (A.�Ա�����=[4] Or A.�Ա����� is NULL)", "") & _
        IIf(mlngPatiType = 1, " And (Nvl(A.���÷�Χ,0) = 0 or A.���÷�Χ =1) ", " And (Nvl(A.���÷�Χ,0) = 0 or A.���÷�Χ =2) ") & vbNewLine & _
        " And (A.����ʱ�� is Null Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Order by A.����"
    strSql = "Select distinct A.Id,A.��ĿID, A.����, A.����,A.����, A.�Ƿ���,A.��������, A.����id,A.�������, " & _
                " Decode(a.����, [6], 1, Decode(A.����,[6],1,decode(a.����,[6],1,NULL))) As ����1ID," & vbNewLine & _
        "                Decode(d.����id, Null, Decode(c.����id, Null, Null, 2), 1) As ����2ID," & vbNewLine & _
        "                Decode(Substr(a.����, 1, Length([6])), [6], 1, Decode(Substr(A.����, 1, Length([6])),[6],1,decode(Substr(a.����, 1, Length([6])),[6],1,NULL))) As ����3ID" & vbNewLine & _
                " From (" & strSql & ") A, ����������� C, ����������� D " & _
                " Where  c.����id(+) = a.Id And d.����id(+) = a.Id And c.����id(+)=[8]  And d.��Աid(+) = [7] " & _
                " Order By ������� desc,����1ID, ����2ID, ����3ID, A.����"
    GetICDElevenSql = strSql
End Function

Private Function MoveCurrRow(ByVal strType As String, ByVal LngRow As Long, ByVal lngWay As Long) As Long
'���ܣ�����ǰ�����ƻ�����һ��
'������lngRow=��ǰ��
'      lngWay=1����һ��,-1����һ��(�൱����һ������һ��)
    Dim lngPreRow As Long, lngNextRow As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim lngUpBegin As Long, lngUpEnd As Long
    Dim lngDownBegin As Long, lngDownEnd As Long
    Dim i As Long, j As Long
    Dim lngMoveRows As Long, blnRedraw As Boolean
    With vsDiagICDEleven
        Call GetRowScope(strType, LngRow, lngBegin, lngEnd)
        If lngWay = 1 Then
            lngPreRow = GetPreRow(lngBegin)
            If lngPreRow = -1 Then Exit Function
            lngDownBegin = lngBegin
            lngDownEnd = lngEnd
            Call GetRowScope(strType, lngPreRow, lngUpBegin, lngUpEnd)
            lngMoveRows = lngDownBegin - lngUpBegin
        Else
            lngNextRow = GetNextRow(lngEnd)
            If lngNextRow = -1 Then Exit Function
            lngUpBegin = lngBegin
            lngUpEnd = lngEnd
            Call GetRowScope(strType, lngNextRow, lngDownBegin, lngDownEnd)
            lngMoveRows = lngDownEnd - lngUpEnd
        End If

        MoveCurrRow = lngMoveRows
        j = 0
        For i = lngDownBegin To lngDownEnd
            .RowPosition(i) = lngUpBegin + j
            j = j + 1
        Next
        
        LngRow = LngRow - lngWay * lngMoveRows
        .Row = LngRow
    End With
End Function

Private Sub GetRowScope(ByVal strType As String, ByVal LngRow As Long, lngBegin As Long, lngEnd As Long)
    Dim i As Long, j As Long, k As Long
    With vsDiagICDEleven
        lngBegin = LngRow: lngEnd = LngRow
        If strType = "��չ��" Or strType = "֤  ��" Then
            j = LngRow
            For i = LngRow To .FixedRows Step -1
                If strType = .TextMatrix(i, Eleven_�������) Then
                    j = i
                Else
                    Exit For
                End If
            Next
            k = LngRow
            For i = .Row To .Rows - 1
                If strType = .TextMatrix(i, Eleven_�������) Then
                    k = i
                Else
                    Exit For
                End If
            Next
        Else
            j = LngRow
            For i = LngRow To .FixedRows Step -1
                If strType = .TextMatrix(i, Eleven_�������) Then
                    j = i
                    Exit For
                End If
            Next
            k = LngRow
            For i = LngRow To .Rows - 1
                If strType <> .TextMatrix(i, Eleven_�������) Then
                    k = i
                Else
                    If i <> LngRow Then
                        Exit For
                    End If
                End If
            Next
            lngBegin = j: lngEnd = k
        End If
    End With
End Sub

Private Function GetPreRow(ByVal LngRow As Long) As Long
'���ܣ�ȡ��һ�����
'���أ�����Ч��ʱ,����-1
    Dim lngTmp As Long, i As Long

    lngTmp = -1
    For i = LngRow - 1 To vsDiagICDEleven.FixedRows Step -1
        lngTmp = i: Exit For
    Next
    GetPreRow = lngTmp
End Function

Private Function GetNextRow(ByVal LngRow As Long) As Long
'���ܣ�ȡ��һ�����
'���أ�����Ч��ʱ,����-1
    Dim lngTmp As Long, i As Long

    lngTmp = -1
    For i = LngRow + 1 To vsDiagICDEleven.Rows - 1
        lngTmp = i: Exit For
    Next
    GetNextRow = lngTmp
End Function

Public Sub InitRsICDEleven(ByRef rsData As ADODB.Recordset)
'���ܣ���ʼ����¼��
    Set rsData = New ADODB.Recordset
    With rsData
        .Fields.Append "�к�", adInteger '��ʼ����ʱ��
        
        .Fields.Append "��ϱ���", adVarChar, 2000 '��������Ŀ¼.����
        .Fields.Append "�������", adVarChar, 4000 '��������Ŀ¼.����
        .Fields.Append "�������", adVarChar, 100 '�������������չ�룬�ַ�����"������"/"��չ��"
        .Fields.Append "IndexEx", adVarChar, 4 ' ¼����� �� 01,0101,0102,0103,02,0201,0202,0203
        .Fields.Append "����ID", adInteger, 100 '��������Ŀ¼.ID
        .Fields.Append "ҽ��IDs", adVarChar, 200
        .Fields.Append "Tag", adVarChar, 4000 '��ҽ,��ҽ
        .Fields.Append "���", adInteger '�������ϵ������ .row
        .Fields.Append "��Ϣ��", adVarChar, 100 '
        
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
End Sub

Public Sub GetRsICD11¼������(ByRef rsData As ADODB.Recordset, ByRef lng����ID As Long, ByRef str����Out As String, ByRef str����Out As String)
'���ܣ��ӹ���¼����ICD11������������Ӧ�ӿ�
'������intType 0-����ǰ�ӹ���1-���غ�ӹ�
    Dim i As Long
    Dim lng���� As Long
    Dim lng���� As Long
    Dim strBackInfo As String
    Dim str���� As String
    Dim rsTmp As ADODB.Recordset
    Dim lng��� As Long
           
    With rsData
        .Filter = 0
        .Sort = "IndexEx"
        lng����ID = Val(!����id & "")
        For i = 1 To rsData.RecordCount
            If !������� & "" = "������" Then
                strBackInfo = strBackInfo & "/" & !�������
                str���� = str���� & "/" & !��ϱ���
            Else
                strBackInfo = strBackInfo & "&" & !�������
                str���� = str���� & "&" & !��ϱ���
            End If
            .MoveNext
        Next
        str����Out = Mid(str����, 2)
        str����Out = Mid(strBackInfo, 2)
        .Filter = 0
        .Sort = "IndexEx"
    End With
End Sub











