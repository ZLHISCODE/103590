VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmFilesSeverConfigure 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "�ļ�����������"
   ClientHeight    =   5220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraOption 
      BorderStyle     =   0  'None
      Height          =   252
      Left            =   120
      TabIndex        =   11
      Top             =   4200
      Width           =   4932
      Begin VB.CheckBox chkSampleServer 
         Caption         =   "����ǰ������ļ��Ƿ���ڣ������ڼ���FTP���ߣ�"
         Height          =   180
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   4932
      End
   End
   Begin VB.PictureBox picBtn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   135
      ScaleHeight     =   336
      ScaleWidth      =   5388
      TabIndex        =   10
      Top             =   180
      Width           =   5385
      Begin VB.CommandButton cmdAdd 
         Caption         =   "����(&A)"
         Height          =   300
         Left            =   0
         TabIndex        =   0
         ToolTipText     =   "����һ���������ռ�������"
         Top             =   0
         Width           =   900
      End
      Begin VB.CommandButton cmdCheck 
         Caption         =   "�����������Լ��(&X)"
         Height          =   300
         Left            =   3240
         TabIndex        =   3
         ToolTipText     =   "����У��������Ƿ������ӳɹ�"
         Top             =   0
         Width           =   2000
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "ɾ��(&D)"
         Enabled         =   0   'False
         Height          =   300
         Left            =   2160
         TabIndex        =   2
         ToolTipText     =   "ɾ��һ����������Ϣ"
         Top             =   0
         Width           =   900
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "�޸�(&S)"
         Enabled         =   0   'False
         Height          =   300
         Left            =   1080
         TabIndex        =   1
         ToolTipText     =   "�޸�һ����������Ϣ"
         Top             =   0
         Width           =   900
      End
   End
   Begin VB.PictureBox picFilter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   165
      ScaleHeight     =   252
      ScaleWidth      =   3996
      TabIndex        =   9
      Top             =   3930
      Width           =   4000
      Begin VB.OptionButton optFilter 
         Caption         =   "ͣ��"
         Height          =   240
         Index           =   2
         Left            =   3195
         TabIndex        =   7
         ToolTipText     =   "��ʾͣ�õķ�����"
         Top             =   0
         Width           =   720
      End
      Begin VB.OptionButton optFilter 
         Caption         =   "����"
         Height          =   240
         Index           =   1
         Left            =   2190
         TabIndex        =   6
         ToolTipText     =   "��ʾ���õķ�����"
         Top             =   0
         Width           =   720
      End
      Begin VB.OptionButton optFilter 
         Caption         =   "ȫ��"
         Height          =   240
         Index           =   0
         Left            =   1080
         TabIndex        =   5
         ToolTipText     =   "��ʾ���з�����"
         Top             =   0
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Label lblFilter 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������б�"
         Height          =   180
         Left            =   0
         TabIndex        =   4
         Top             =   15
         Width           =   900
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfMain 
      Height          =   2835
      Left            =   120
      TabIndex        =   8
      Top             =   700
      Width           =   6870
      _cx             =   12118
      _cy             =   5001
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFilesSeverCongifure.frx":0000
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
      ExplorerBar     =   5
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
End
Attribute VB_Name = "frmFilesSeverConfigure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrSelectSeverNum As String '��λ�з��������
Private mblnFirstAdd As Boolean '��һ�����ӷ�������Ҫ����ΪĬ�Ϸ�����
Public blnRefreshData As Boolean '�����л�ˢ���жϱ�־

Private Enum ServerListCols
    Col_��� = 0 '״ֵ̬ 0-���� 1-��������ȱʧ(�����ļ��ز�����) 2-�����ļ������� 3-������� 4-���浫�����ϴ� 5-�Ѿ��ϴ�
    Col_���� = 1
    Col_������״̬ = 2 '���� or ͣ��
    Col_������·�� = 3
    Col_�û��� = 4
    Col_���� = 5
    Col_�˿� = 6
    Col_�Ƿ����� = 7
    Col_�Ƿ�ȱʡ = 8
    Col_�Ƿ��ռ� = 9
    Col_�ռ����� = 10
    Col_����� = 11
    Col_�������б����� = 12
End Enum

Private Const SS_ͣ�� = "ͣ��"
Private Const SS_���� = "����"
Private Const ST_FTP = "FTP"
Private Const ST_���� = "����"

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�
End Sub


Private Sub chkSampleServer_Click()
    Dim strSQL As String
    
    If chkSampleServer.Tag <> "" Then
        strSQL = "Update Zlreginfo Set ���� = '" & chkSampleServer.value & "' Where ��Ŀ = 'FTP������ļ�����'"
        gcnOracle.Execute strSQL, , adCmdText
    End If
End Sub

Private Sub cmdAdd_Click()
    Dim frmEdit As New frmFilesSeverEdit
    
    If frmEdit.ShowMe(0, mblnFirstAdd) Then
        LoadSeverListData
        vsfMain.Row = vsfMain.Rows - 1
        vsfMain.SetFocus
    End If
End Sub

Private Sub cmdCheck_Click()
    Dim strSeverAddress As String
    Dim strUser As String
    Dim strPassword As String
    Dim strPort As String
    Dim lngRowsCount As Long
    Dim strInformation As String
    Dim i As Long
    
    With vsfMain
        If .Rows < .FixedRows Then Exit Sub
        lngRowsCount = .Rows - 1
        .Row = 0
        cmdDel.Enabled = False: cmdModify.Enabled = False
        For i = .FixedRows To lngRowsCount
            ShowFlash "���ڼ��" & .TextMatrix(i, Col_���) & "��: " & .TextMatrix(i, Col_������·��), i / (lngRowsCount), Me, True
            DoEvents
            If .TextMatrix(i, Col_������״̬) = SS_ͣ�� Then
                .TextMatrix(i, Col_�����) = "�����ã�" & "�÷���������ͣ��״̬�������ú�������У�顣"
            Else
                strSeverAddress = Trim(.TextMatrix(i, Col_������·��))
                strUser = Trim(.TextMatrix(i, Col_�û���))
                strPassword = Trim(.Cell(flexcpData, i, Col_����))
                strPort = Trim(.TextMatrix(i, Col_�˿�))
                
                If Trim(.TextMatrix(i, Col_����)) = ST_FTP Then
                    If CheckFTPServer(strSeverAddress, strUser, strPassword, strPort, strInformation) = False Then
                        .TextMatrix(i, Col_�����) = "�����ã�" & strInformation
                    Else
                        .TextMatrix(i, Col_�����) = "����"
                    End If
                Else
                    If CheckFileServer(strSeverAddress, strUser, strPassword, strInformation) = False Then
                        .TextMatrix(i, Col_�����) = "�����ã�" & strInformation
                    Else
                        .TextMatrix(i, Col_�����) = "����"
                    End If
                End If
            End If
        Next
        ShowFlash ("")
    End With
End Sub

Private Sub cmdDel_Click()
    Dim strSQL As String
    If vsfMain.TextMatrix(vsfMain.Row, Col_�Ƿ�ȱʡ) <> "" Then
        MsgBox vsfMain.TextMatrix(vsfMain.Row, Col_���) & " ��" & vsfMain.TextMatrix(vsfMain.Row, Col_����) & "������Ϊȱʡ����������ɾ�������л�ȱʡ��������ɾ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("ȷ��Ҫɾ�� " & vsfMain.TextMatrix(vsfMain.Row, Col_���) & " ��" & vsfMain.TextMatrix(vsfMain.Row, Col_����) & "��������", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
        strSQL = "Zl_Zlupgradeserver_Delete('" & vsfMain.TextMatrix(vsfMain.Row, Col_���) & "')"
        Call ExecuteProcedure(strSQL, Me.Caption)
        strSQL = "update ZLClients set �����ļ������� = null where �����ļ������� = " & vsfMain.TextMatrix(vsfMain.Row, Col_���)
        gcnOracle.Execute strSQL
'        Load frmUpgradeManage
        
        Call LoadSeverListData
        vsfMain.SetFocus
    End If
End Sub

Private Sub cmdModify_Click()
    Dim frmEdit As New frmFilesSeverEdit

    If frmEdit.ShowMe(1, mblnFirstAdd, Nvl(vsfMain.TextMatrix(vsfMain.Row, 0), "")) Then
        LoadSeverListData
    End If
    
End Sub

Private Sub Form_Load()
    '���ط�����������Ϣ
    If TransData = False Then MsgBox "�ɰ汾����������ת��ʧ�ܣ�����ϵ������Ա��", vbInformation, gstrSysName
'    Call RefreshData
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fraOption.Top = Me.ScaleHeight - fraOption.Height - 60
    vsfMain.Move 50, 650, Me.ScaleWidth - 120, fraOption.Top - vsfMain.Top - 120
    picBtn.Move 50, 210
    If Me.ScaleWidth < 8000 Then picFilter.Visible = False
    If Me.ScaleWidth >= 8000 Then picFilter.Visible = True
    picFilter.Move Me.ScaleWidth - picFilter.Width, 280
End Sub

Public Sub SetMenu()
    frmMDIMain.stbThis.Panels(2).Text = "�б��й���ʾ��" & vsfMain.Rows - 1 & "�����ݡ�"
End Sub

Public Sub LoadSeverListData(Optional ByVal strFilter As String, Optional ByVal strLocationName As String)
    Dim i, j As Long
    Dim strSQL       As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngLocationRow As Long
    
    On Error GoTo errH

    mblnFirstAdd = True
    strSQL = "Select ���� As ʹ�ü���ftp���� From Zlreginfo Where ��Ŀ =[1]"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, "FTP������ļ�����")
    If rsTemp.EOF Then
        chkSampleServer.value = 0
        strSQL = "Insert Into Zlreginfo (��Ŀ, ����) Values ('FTP������ļ�����', '0')"
        gcnOracle.Execute strSQL, , adCmdText
    Else
        chkSampleServer.value = Val(rsTemp!ʹ�ü���ftp���� & "")
    End If
    chkSampleServer.Tag = "�����Ѿ�����"
    With vsfMain

        If .Row < .FixedRows Then .Row = 0
        lngLocationRow = .Row
    
        .Redraw = flexRDNone
        .Rows = .FixedRows
'        .Clear
'        .Cols = Col_�������б�����

'        strSQL = "select ���,����,λ��,�û���,����,�˿�,�Ƿ�����,�Ƿ�ȱʡ,�Ƿ��ռ�,�ռ����� from ZLUpgradeServer order by ���"
        strSQL = "select ���,����,λ��,�û���,����,�˿�,�Ƿ�����,�Ƿ�ȱʡ from ZLUpgradeServer order by ���"
        Call OpenRecordset(rsTemp, strSQL, Me.Caption)

        '��������
        .Rows = rsTemp.RecordCount + 1
        i = .FixedRows
        Do Until rsTemp.EOF
        
            .TextMatrix(i, Col_���) = Nvl(rsTemp.Fields("���"), "")
            .TextMatrix(i, Col_����) = IIf(Nvl(rsTemp.Fields("����"), "") = "1", ST_FTP, ST_����)
            .TextMatrix(i, Col_������·��) = Nvl(rsTemp.Fields("λ��"), "")
            .TextMatrix(i, Col_�û���) = Nvl(rsTemp.Fields("�û���"), "")
            .TextMatrix(i, Col_����) = "***"
            .Cell(flexcpData, i, Col_����) = Decipher(Nvl(rsTemp.Fields("����"), ""))
            .TextMatrix(i, Col_�˿�) = Nvl(rsTemp.Fields("�˿�"), "")
            .Cell(flexcpBackColor, i, Col_�Ƿ�����, i, Col_�Ƿ��ռ�) = RGB(210, 240, 255) 'RGB(247, 247, 247)
            .TextMatrix(i, Col_�Ƿ�����) = IIf(Nvl(rsTemp.Fields("�Ƿ�����"), "") = "1", "��", "")
            .TextMatrix(i, Col_�Ƿ�ȱʡ) = IIf(Nvl(rsTemp.Fields("�Ƿ�ȱʡ"), "") = "1", "��", "")
            If .TextMatrix(i, Col_�Ƿ�ȱʡ) = "��" Then
                mblnFirstAdd = False
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlue
            End If
'            .TextMatrix(i, Col_�Ƿ��ռ�) = IIf(Nvl(rsTemp.Fields("�Ƿ��ռ�"), "") = "1", "��", "")
            .TextMatrix(i, Col_�����) = ""
'            .TextMatrix(i, Col_�ռ�����) = Nvl(rsTemp.Fields("�ռ�����"), "")
            
            If .TextMatrix(i, Col_�Ƿ�����) = "" And .TextMatrix(i, Col_�Ƿ�ȱʡ) = "" And .TextMatrix(i, Col_�Ƿ��ռ�) = "" Then
                .TextMatrix(i, Col_������״̬) = SS_ͣ��
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbGrayText
            Else
                .Cell(flexcpText, i, Col_������״̬) = SS_����
            End If

            rsTemp.MoveNext
            i = i + 1
        Loop
        
        'ѡ�п���
        .FocusRect = flexFocusSolid
        '���һ���Զ��п�
        .ExtendLastCol = True
        '�����������
        .ScrollTrack = True
        '�Զ�����
        .WordWrap = True
        '�и�����
        .RowHeightMin = 300
        .RowHeightMax = 300
        '����������
        .ColWidthMax = 7000
        '�Զ���Ӧ�иߡ��п�
        .AutoSizeMode = flexAutoSizeRowHeight
        .SelectionMode = flexSelectionListBox
        .AllowBigSelection = False
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False
        
        If lngLocationRow > .Rows - 1 Then lngLocationRow = .Rows - 1
        .Row = lngLocationRow
        .Redraw = flexRDBuffered
        
        Call SetMenu
        
    End With
    Exit Sub
errH:
    Call MsgBox("�������б����ش���", vbInformation, gstrSysName)
    If False Then
        Resume
    End If
End Sub

Private Sub optFilter_Click(Index As Integer)
Dim i As Long
    With vsfMain
        If .Rows < 1 Then Exit Sub
        .Redraw = flexRDNone
        For i = 1 To .Rows - 1
            Select Case Index
            Case 0
                .RowHidden(i) = False
            Case 1
                .RowHidden(i) = .TextMatrix(i, Col_������״̬) = SS_ͣ��
            Case 2
                .RowHidden(i) = .TextMatrix(i, Col_������״̬) = SS_����
            End Select
        Next
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub vsfMain_AfterSort(ByVal Col As Long, Order As Integer)
    vsfMain.Row = vsfMain.FindRow(mstrSelectSeverNum, , Col_���)
End Sub

Private Sub vsfMain_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        If Row = 0 Then Cancel = True
End Sub

Private Sub vsfMain_DblClick()
    Dim strIsUpgrade As String
    Dim strIsCheck As String
    Dim strIsCollect As String
    Dim strFilesType As String
    Dim strSQL As String
    On Error GoTo errHand
    
    With vsfMain
        If .MouseRow <> .Row Then Exit Sub
        strIsUpgrade = IIf(.TextMatrix(.Row, Col_�Ƿ�����) = "��", "1", "0")
        strIsCheck = IIf(.TextMatrix(.Row, Col_�Ƿ�ȱʡ) = "��", "1", "0")
        strIsCollect = IIf(.TextMatrix(.Row, Col_�Ƿ��ռ�) = "��", "1", "0")
        strFilesType = .TextMatrix(.Row, Col_�ռ�����)

        Select Case .ColSel
        Case Col_�Ƿ�����
            If strIsCheck = "1" Then
                Call MsgBox("ѡ�б�� " & .TextMatrix(.Row, Col_���) & " ������Ϊȱʡ���������뱣֤������һ��ȱʡ�������� ", vbInformation, gstrSysName)
                Exit Sub
            ElseIf strIsCollect = "1" Then
                If MsgBox("ѡ�б�� " & .TextMatrix(.Row, Col_���) & " ������Ϊ�ռ����������Ƿ�Ҫ�л�Ϊ������������ ", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                    If mblnFirstAdd = True Then
                        strFilesType = ""
                        strIsUpgrade = "1"
                        strIsCheck = "1"
                        strIsCollect = "0"
                    Else
                        strFilesType = ""
                        strIsUpgrade = "1"
                        strIsCollect = "0"
                    End If
                End If
            ElseIf mblnFirstAdd = True Then
                strIsUpgrade = "1"
                strIsCheck = "1"
            Else
                If strIsUpgrade = "1" Then
                    If MsgBox("�Ƿ�Ҫȡ����������������ȡ���󽫻���������ù��÷�����Ϊ�����������Ŀͻ���", vbOKCancel, gstrSysName) = vbOK Then
                        strIsUpgrade = "0"
                    End If
                Else
                    strIsUpgrade = "1"
                End If
            End If
        Case Col_�Ƿ�ȱʡ
            If strIsCheck = "1" Then
                Call MsgBox("ѡ�б�� " & .TextMatrix(.Row, Col_���) & " ������ΪĬ�Ϸ��������뱣֤������һ��Ĭ�Ϸ����� ", vbInformation, gstrSysName)
                Exit Sub
            ElseIf strIsCollect = "1" Then
                If MsgBox("ѡ�б�� " & .TextMatrix(.Row, Col_���) & " ������Ϊ�ռ����������Ƿ�Ҫ�л�Ϊ����������������Ϊȱʡ�������� ", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                    strFilesType = ""
                    strIsUpgrade = "1"
                    strIsCheck = "1"
                    strIsCollect = "0"
                End If
            ElseIf strIsUpgrade = "0" Then
                If MsgBox("ѡ�б�� " & .TextMatrix(.Row, Col_���) & " ������Ϊͣ��״̬���Ƿ�Ҫ���ø÷�����������Ϊȱʡ�������� ", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                    strFilesType = ""
                    strIsUpgrade = "1"
                    strIsCheck = "1"
                    strIsCollect = "0"
                End If
            Else
                strIsCheck = IIf(strIsCheck = "0", "1", "0")
            End If
        Case Col_�Ƿ��ռ�
            If strIsCheck = "1" Then
                Call MsgBox("ѡ�б�� " & .TextMatrix(.Row, Col_���) & " ������ΪȱʡĬ������������,�����л�Ϊ�ռ���������", vbInformation, gstrSysName)
                Exit Sub
            ElseIf strIsUpgrade = "1" Then
                If MsgBox("ѡ�б�� " & .TextMatrix(.Row, Col_���) & " ������Ϊ�������������Ƿ�Ҫ�л�Ϊ�ռ��������� ", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                    If .TextMatrix(.Row, Col_�ռ�����) = "" Then strFilesType = "Log"
                    strIsUpgrade = "0"
                    strIsCheck = "0"
                    strIsCollect = "1"
                End If
            Else
                strIsCollect = IIf(strIsCollect = "1", "0", "1")
            End If
        Case Else
            Exit Sub
        End Select

        strSQL = "Zl_Zlupgradeserver_Update('" & .TextMatrix(.Row, Col_���) & "','','','','','','" & strIsUpgrade & "','" & strIsCheck & "','" & strIsCollect & "','" & strFilesType & "','" & 1 & "')"
        Call ExecuteProcedure(strSQL, Me.Caption)
        
        If strIsUpgrade = "0" Then
            strSQL = "update ZLClients set �����ļ������� = null where �����ļ������� = " & .TextMatrix(.Row, Col_���)
            gcnOracle.Execute strSQL
        End If
        
        If strIsCheck = "1" Then
            strSQL = "ZLReginfo_DefaultServer('" & IIf(.TextMatrix(.Row, Col_����) = ST_����, "0", "1") & "','" & Trim(.TextMatrix(.Row, Col_������·��)) & "','" & Trim(.TextMatrix(.Row, Col_�û���)) & "','" & Trim(.Cell(flexcpData, .Row, Col_����)) & "','" & Trim(.TextMatrix(.Row, Col_�˿�)) & "')"
            Call ExecuteProcedure(strSQL, Me.Caption)
        End If

        LoadSeverListData
        optFilter.Item(0).value = True
    End With
    
    Exit Sub
errHand:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub vsfMain_RowColChange()
    mstrSelectSeverNum = vsfMain.TextMatrix(vsfMain.Row, Col_���)
    If vsfMain.Row > 0 Then cmdModify.Enabled = True: cmdDel.Enabled = True: cmdCheck.Enabled = True
End Sub

Private Function CheckFTPServer(ByVal strIp As String, ByVal strUser As String, ByVal strPass As String, ByVal strPort As String, Optional ByRef strError As String) As Boolean
    '-----------------------------------------------------------------------------
    '����:��鵱ǰ��FTP�������Ƿ���ȷ
    '����:��ǰ���ļ��������ĸ�����ȷ,����true,���򷵻�False
    '����:����ԭ
    '����:2016/07/05
    'strIp - FTP��ַ
    'strUser - �û���
    'strPass - ����
    'strPort - �˿�
    '-----------------------------------------------------------------------------
    On Error GoTo errHand:
    
    If strIp = "" Or strUser = "" Or strPass = "" Or strPort = "" Then
        CheckFTPServer = False
        Exit Function
    End If
    
    If IsFtpServer(Trim(strIp), Trim(strUser), Trim(strPass), Trim(strPort)) Then
        CheckFTPServer = True
        strError = "���ӳɹ�"
    Else
        CheckFTPServer = False
        strError = "������������������������FTP����������"
    End If
    CancelFtpServer
    Exit Function
    
errHand:
        MsgBox err.Description, vbInformation, gstrSysName
End Function


Private Function CheckFileServer(ByVal strAddress As String, ByVal strUser As String, ByVal strPass As String, Optional ByRef strError As String) As Boolean
    '-----------------------------------------------------------------------------
    '����:��鵱ǰ���ļ��������Ƿ���ȷ
    '����:��ǰ���ļ��������ĸ�����ȷ,����true,���򷵻�False
    '����:����ԭ
    '����:2016/07/05
    'strAddress - ��ַ
    'strUser - �û�
    'strPass - ����
    '-----------------------------------------------------------------------------
    Dim typOfStruct As OFSTRUCT

    On Error GoTo errHand:
    
    If strAddress = "" Or strUser = "" Or strPass = "" Then
        CheckFileServer = False
        Exit Function
    End If
    
    If IsNetServer(Trim(strAddress), Trim(strUser), Trim(strPass)) = False Then
        strError = "�����ļ���ָ��Ŀ¼������,����������"
        CheckFileServer = False
    Else
        strError = "���ӳɹ�"
        CheckFileServer = True
    End If
    Call CancelNetServer(Trim(strAddress))
    
    Exit Function
errHand:
        MsgBox err.Description, vbInformation, gstrSysName
End Function

Private Function FindFile(ByVal strFileName As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------
    '--����:����ָ�����ļ����ļ��Ƿ����
    '--����: ������ڴ��ļ�ΪTrue,����ΪFlase
    '------------------------------------------------------------------------------------------------------------------------------------
    Dim typOfStruct As OFSTRUCT
    
    On Error Resume Next
    FindFile = False
    If Len(strFileName) > 0 Then
        apiOpenFile strFileName, typOfStruct, OF_EXIST
        FindFile = typOfStruct.nErrCode <> 2
    End If
End Function

Private Function TransData() As Boolean
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim strCollectType As String
    Dim intNoNum As Integer
    Dim intSeverNum As Integer
    Dim intUpType As Integer
    Dim i As Long
    
    On Error GoTo errH
    strSQL = "select * from zlupgradeserver"
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
    If Not rsTemp.EOF Then TransData = True: Exit Function
'    If MsgBox("�Ƿ񽫾ɰ汾���ù��ķ���������ת�����°汾���������ã�", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then TransData = True: Exit Function
    
    '��տͻ����������õ�����������
    strSQL = "update zlclients set �����ļ������� = null,���������� = null,FTP������=null"
    gcnOracle.Execute strSQL
    
    '����FTP 0-N ����������
    strSQL = "select max(nvl(replace(��Ŀ,'FTP������',''),'-1')) as �������� from zlreginfo where ��Ŀ like 'FTP������%'"
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
        
    intSeverNum = Nvl(rsTemp.Fields("��������"), -1)
        
    If intSeverNum >= 0 Then
        For i = 0 To intSeverNum
            strSQL = "Select Max(Decode(��Ŀ, 'FTP�˿�', ����, '')) FTP�˿�, Max(Decode(��Ŀ, 'FTP������', ����, '')) FTP������," & vbNewLine & _
                        "       Max(Decode(��Ŀ, 'FTP����', ����, '')) FTP����, Max(Decode(��Ŀ, 'FTP�û�', ����, '')) FTP�û�" & vbNewLine & _
                        "From (Select Substr(��Ŀ, Length(��Ŀ), 1) ID, Substr(��Ŀ, 1, Length(��Ŀ) - 1) ��Ŀ, ����" & vbNewLine & _
                        "       From zlRegInfo" & vbNewLine & _
                        "       Where ��Ŀ = 'FTP������" & i & "' or ��Ŀ = 'FTP�û�" & i & "' or ��Ŀ = 'FTP����" & i & "' or ��Ŀ = 'FTP�˿�" & i & "')" & vbNewLine & _
                        "Group By ID"
            Call OpenRecordset(rsTemp, strSQL, Me.Caption)
                
            If rsTemp.EOF = False Then
                If Nvl(rsTemp.Fields("FTP������"), "") <> "" And Nvl(rsTemp.Fields("FTP�û�"), "") <> "" And Nvl(rsTemp.Fields("FTP����"), "") <> "" And Nvl(rsTemp.Fields("FTP�˿�"), "") <> "" Then
                    intNoNum = intNoNum + 1
                    strSQL = "Zl_Zlupgradeserver_Insert('" & intNoNum & "','" & 1 & "','" & rsTemp.Fields("FTP������") & "','" & rsTemp.Fields("FTP�û�") & "','" & Cipher(rsTemp.Fields("FTP����")) & "','" & rsTemp.Fields("FTP�˿�") & "','" & 1 & "','" & 0 & "','" & 0 & "','')"
                    Call ExecuteProcedure(strSQL, Me.Caption)
                End If
            Else
                
            End If
        Next
    End If
        
    '����FTP������������
    strSQL = "Select Max(Decode(��Ŀ, 'FTP�˿�', ����, '')) FTP�˿�, Max(Decode(��Ŀ, 'FTP������', ����, '')) FTP������," & vbNewLine & _
                "       Max(Decode(��Ŀ, 'FTP����', ����, '')) FTP����, Max(Decode(��Ŀ, 'FTP�û�', ����, '')) FTP�û�" & vbNewLine & _
                "From (Select Substr(��Ŀ, Length(��Ŀ), 1) ID, Substr(��Ŀ, 1, Length(��Ŀ) - 1) ��Ŀ, ����" & vbNewLine & _
                "       From zlRegInfo" & vbNewLine & _
                "       Where (��Ŀ = 'FTP������' or ��Ŀ = 'FTP�û�' or ��Ŀ = 'FTP����' or ��Ŀ = 'FTP�˿�') And Not ���� Is Null)" & vbNewLine & _
                "Group By ID"
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
            
    If rsTemp.EOF = False Then
        If Nvl(rsTemp.Fields("FTP������"), "") <> "" And Nvl(rsTemp.Fields("FTP�û�"), "") <> "" And Nvl(rsTemp.Fields("FTP����"), "") <> "" And Nvl(rsTemp.Fields("FTP�˿�"), "") <> "" Then
            intNoNum = intNoNum + 1
            strSQL = "Zl_Zlupgradeserver_Insert('" & intNoNum & "','" & 1 & "','" & rsTemp.Fields("FTP������") & "','" & rsTemp.Fields("FTP�û�") & "','" & Cipher(rsTemp.Fields("FTP����")) & "','" & rsTemp.Fields("FTP�˿�") & "','" & 1 & "','" & 0 & "','" & 0 & "','')"
            Call ExecuteProcedure(strSQL, Me.Caption)
        End If
    Else
    End If

    '���� ���� 0-N ����������
    strSQL = "select max(nvl(replace(��Ŀ,'������Ŀ¼',''),'-1')) as �������� from zlreginfo where ��Ŀ like '������Ŀ¼%'"
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
    intSeverNum = Nvl(rsTemp.Fields("��������"), -1)
        
    If intSeverNum >= 0 Then
        For i = 0 To intSeverNum
            strSQL = "Select Max(Decode(��Ŀ, '������Ŀ¼', ����, '')) ������Ŀ¼, Max(Decode(��Ŀ, '�����û�', ����, '')) �����û�," & vbNewLine & _
                        "       Max(Decode(��Ŀ, '��������', ����, '')) ��������" & vbNewLine & _
                        "From (Select Substr(��Ŀ, Length(��Ŀ), 1) ID, Substr(��Ŀ, 1, Length(��Ŀ) - 1) ��Ŀ, ����" & vbNewLine & _
                        "       From zlRegInfo" & vbNewLine & _
                        "       Where ��Ŀ = '������Ŀ¼" & i & "' Or ��Ŀ = '�����û�" & i & "' Or ��Ŀ = '��������" & i & "')" & vbNewLine & _
                        "Group By ID"
            Call OpenRecordset(rsTemp, strSQL, Me.Caption)
                
            If rsTemp.EOF = False Then
                If Nvl(rsTemp.Fields("������Ŀ¼"), "") <> "" And Nvl(rsTemp.Fields("�����û�"), "") <> "" And Nvl(rsTemp.Fields("��������"), "") <> "" Then
                    intNoNum = intNoNum + 1
                    strSQL = "Zl_Zlupgradeserver_Insert('" & intNoNum & "','" & 0 & "','" & rsTemp.Fields("������Ŀ¼") & "','" & rsTemp.Fields("�����û�") & "','" & Cipher(rsTemp.Fields("��������")) & "','','" & 1 & "','" & 0 & "','" & 0 & "','')"
                    Call ExecuteProcedure(strSQL, Me.Caption)
                End If
            End If
        Next
    End If
    
    '���� ���� ������������
    strSQL = "Select Max(Decode(��Ŀ, '������Ŀ¼', ����, '')) ������Ŀ¼, Max(Decode(��Ŀ, '�����û�', ����, '')) �����û�," & vbNewLine & _
                "       Max(Decode(��Ŀ, '��������', ����, '')) ��������" & vbNewLine & _
                "From (Select 1 ID, ��Ŀ, ����" & vbNewLine & _
                "       From zlRegInfo" & vbNewLine & _
                "       Where (��Ŀ = '������Ŀ¼' Or ��Ŀ = '�����û�' Or ��Ŀ = '��������') And Not ���� Is Null)" & vbNewLine & _
                "Group By ID"
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
            
    If rsTemp.EOF = False Then
        If Nvl(rsTemp.Fields("������Ŀ¼"), "") <> "" And Nvl(rsTemp.Fields("�����û�"), "") <> "" And Nvl(rsTemp.Fields("��������"), "") <> "" Then
            intNoNum = intNoNum + 1
            strSQL = "Zl_Zlupgradeserver_Insert('" & intNoNum & "','" & 0 & "','" & rsTemp.Fields("������Ŀ¼") & "','" & rsTemp.Fields("�����û�") & "','" & Cipher(rsTemp.Fields("��������")) & "','','" & 1 & "','" & 0 & "','" & 0 & "','')"
            Call ExecuteProcedure(strSQL, Me.Caption)
        End If
    End If

    '�ռ�����������
'    strSQL = "select ���� as �ռ����� from zlreginfo where ��Ŀ = '�ռ�����'"
'    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
'    strCollectType = Nvl(rsTemp.Fields("�ռ�����"), "")
'
'    strSQL = "Select Max(Decode(��Ŀ, '�ռ�Ŀ¼S', ����, '')) �ռ�Ŀ¼, Max(Decode(��Ŀ, '�����û�S', ����, '')) �����û�," & vbNewLine & _
'                "       Max(Decode(��Ŀ, '��������S', ����, '')) ��������, Max(Decode(��Ŀ, '�ռ�����', ����, '')) �ռ�����" & vbNewLine & _
'                "From (Select 1 As ID, ��Ŀ, ����" & vbNewLine & _
'                "       From zlRegInfo" & vbNewLine & _
'                "       Where (��Ŀ = '�ռ�Ŀ¼S' Or ��Ŀ = '�����û�S' Or ��Ŀ = '��������S') And Not ���� Is Null)" & vbNewLine & _
'                "Group By ID"
'    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
'
'    If rsTemp.EOF = False Then
'        intNoNum = intNoNum + 1
'        strSQL = "Zl_Zlupgradeserver_Insert('" & intNoNum & "','" & 0 & "','" & Nvl(rsTemp.Fields("�����û�"), "��Ч") & "','" & Nvl(rsTemp.Fields("��������")) & "','" & Cipher(Nvl(rsTemp.Fields("��������"), "��Ч")) & "','','" & 0 & "','" & 0 & "','" & 1 & "','" & strCollectType & "')"
'        Call ExecuteProcedure(strSQL, Me.Caption)
'    Else
'    End If
'
'    strSQL = "Select Max(Decode(��Ŀ, '�ռ�Ŀ¼F', ����, '')) �ռ�Ŀ¼, Max(Decode(��Ŀ, '�����û�F', ����, '')) �����û�," & vbNewLine & _
'                "       Max(Decode(��Ŀ, '��������F', ����, '')) ��������, Max(Decode(��Ŀ, '�ռ�����', ����, '')) �ռ�����" & vbNewLine & _
'                "From (Select 1 As ID, ��Ŀ, ����" & vbNewLine & _
'                "       From zlRegInfo" & vbNewLine & _
'                "       Where (��Ŀ = '�ռ�Ŀ¼F' Or ��Ŀ = '�����û�F' Or ��Ŀ = '��������F') And Not ���� Is Null)" & vbNewLine & _
'                "Group By ID"
'    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
'
'    If rsTemp.EOF = False Then
'        intNoNum = intNoNum + 1
'        strSQL = "Zl_Zlupgradeserver_Insert('" & intNoNum & "','" & 1 & "','" & Nvl(rsTemp.Fields("�����û�"), "��Ч") & "','" & Nvl(rsTemp.Fields("��������")) & "','" & Cipher(Nvl(rsTemp.Fields("��������"), "��Ч")) & "','','" & 0 & "','" & 0 & "','" & 1 & "','" & strCollectType & "')"
'        Call ExecuteProcedure(strSQL, Me.Caption)
'    Else
'    End If
        
    'ɾ��������
'    strSQL = "delete from zlreginfo where ��Ŀ like 'FTP%'or ��Ŀ like '����%' or ��Ŀ like '������Ŀ¼%'or ��Ŀ like '�ռ�Ŀ¼%'"
'    gcnOracle.Execute strSQL
        
    '����Ĭ�Ϸ�����
    If intNoNum > 0 Then
        strSQL = "select max(����) as �������� from zlreginfo where ��Ŀ = '��������'"
        Call OpenRecordset(rsTemp, strSQL, Me.Caption)
        
        intUpType = Nvl(rsTemp.Fields("��������"), 0)

        strSQL = "select ���,����,λ��,�û���,����,�˿� from zlupgradeserver where ��� = (select min(���) from (select ��� from zlupgradeserver where ���� = " & intUpType & "))"
        Call OpenRecordset(rsTemp, strSQL, Me.Caption)
            
        If rsTemp.EOF Then
            strSQL = "select ���,����,λ��,�û���,����,�˿� from zlupgradeserver where ��� = 1"
            Call OpenRecordset(rsTemp, strSQL, Me.Caption)
        End If
        
        strSQL = "Zl_Zlupgradeserver_Update('" & rsTemp.Fields("���") & "','','','','','','" & 1 & "','" & 1 & "','" & 0 & "','','" & 1 & "')"
        Call ExecuteProcedure(strSQL, Me.Caption)
        strSQL = "ZLReginfo_DefaultServer('" & rsTemp.Fields("����") & "','" & rsTemp.Fields("λ��") & "','" & rsTemp.Fields("�û���") & "','" & Decipher(rsTemp.Fields("����")) & "','" & rsTemp.Fields("�˿�") & "')"
        Call ExecuteProcedure(strSQL, Me.Caption)
    End If
    strSQL = "Select ���� As ʹ�ü���ftp���� From Zlreginfo Where ��Ŀ =[1]"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, "FTP������ļ�����")
    If rsTemp.EOF Then
        chkSampleServer.value = 0
        strSQL = "Insert Into Zlreginfo (��Ŀ, ����) Values ('FTP������ļ�����', '0')"
        gcnOracle.Execute strSQL, , adCmdText
    Else
        chkSampleServer.value = Val(rsTemp!ʹ�ü���ftp���� & "")
    End If
    chkSampleServer.Tag = "�����Ѿ�����"
    TransData = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    TransData = False
    If False Then
        Resume
    End If
End Function

Public Sub RefreshData()
    Call LoadSeverListData
End Sub
