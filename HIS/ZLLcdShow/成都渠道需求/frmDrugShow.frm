VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmDrugShow 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8736
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11640
   ForeColor       =   &H80000001&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8736
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer timerPage 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   7500
      Top             =   8160
   End
   Begin VB.Timer timerLCD 
      Interval        =   10000
      Left            =   9090
      Top             =   8190
   End
   Begin VB.Timer TimerCall 
      Interval        =   1000
      Left            =   90
      Top             =   8160
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgCallingData 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _cx             =   20981
      _cy             =   12938
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   0
      ForeColor       =   65280
      BackColorFixed  =   0
      ForeColorFixed  =   0
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   0
      BackColorAlternate=   0
      GridColor       =   65280
      GridColorFixed  =   65280
      TreeColor       =   -2147483633
      FloodColor      =   0
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDrugShow.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
   Begin VB.Label lblmsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   7440
      Width           =   11655
   End
End
Attribute VB_Name = "frmDrugShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mintCols As Integer
Private mstrWins As String
Private mintRows As Integer
Private mrsData() As Recordset
Private mrsCallingData As Recordset
Private mrsPreparingData() As Recordset
Private mrsTimeoutData() As Recordset
Private mRowRec As Integer
Private mlngҩ��ID As Long
Private mIntCallRow As Integer
Private mIntPraRow As Integer
Private mIntCallCol As Integer
Private mIntPraCol As Integer
Private mIntTimeoutCol As Integer
Private mbln��ҩ As Boolean
Private mbln��ҩȷ�� As Boolean
Private mIntSendPages() As Integer

Private Type Type_para
    bln��������ʾģʽ As Boolean             '������ʾģʽ�������壺�ര��
    Str���� As String
    dblLeft As Double
    dblTop As Double
    dblWidth As Double
    dblHeight As Double
    
    lng������������ɫ As Long
    
    bln��ʾ����ҩ As Boolean
    bln��ʾ��ҩ��� As Boolean
    int����ҩ���� As Integer
    int����ҩ���� As Integer
    int����ҩ���� As Integer
    lng����ҩ������ɫ As Long
    
    bln��ʾ����ҩ As Boolean
    int����ҩ���� As Integer
    int����ҩ���� As Integer
    int����ҩ���� As Integer
    lng����ҩ������ɫ As Long
    
    bln��ʾ�ѹ��� As Boolean
    int�ѹ������� As Integer
    int�ѹ������� As Integer
    int�ѹ������� As Integer
    lng�ѹ���������ɫ As Long
    lng��ǰ����ҳ�� As Long
    
    bln��ʾ���� As Boolean
    lng����������ɫ As Long
    
    bln��ʾ�������� As Boolean
    lng��������������ɫ As Long
    
    
    intRowPeople  As Integer
    intPage As Integer
    intRefTime As Integer
    intTimeout As Integer
    
    str��ʾ���� As String
End Type

Private mType_para As Type_para

Public Sub SetFacePostion()
'************************************************************************************
'
'���ý������ʾλ��
'
'************************************************************************************
    Dim strReg As String
    
    On Error GoTo errHandle
        
    '��ע����У���ȡ��ʾ����
    strReg = "����ģ��\ҩ���Ŷӽк�\Һ������"
    
    '������ʾ����
    Me.Left = GetSetting("ZLSOFT", strReg, "��", "1024") * Screen.TwipsPerPixelX
    Me.Top = GetSetting("ZLSOFT", strReg, "��", "0") * Screen.TwipsPerPixelY
    Me.Width = GetSetting("ZLSOFT", strReg, "���", "1024") * Screen.TwipsPerPixelX
    Me.Height = GetSetting("ZLSOFT", strReg, "�߶�", "768") * Screen.TwipsPerPixelY
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************SetFacePostion*************")
        Call SaveErrLog("������ʱ��:" & gobjDatabase.Currentdate)
        Call SaveErrLog("������Ϣ:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub
Private Sub LoadPara()
    Dim strReg As String
    Dim i As Integer
    Dim strWin As String
    Dim rsWin As New Recordset
    Dim strWins_temp As String, strSql As String
    On Error GoTo errHandle
    strReg = "����ģ��\ҩ���Ŷӽк�\Һ������"
    
    With mType_para
        .bln��������ʾģʽ = (Val(GetSetting("ZLSOFT", strReg, "����ģʽ", "0")) = 0)
        
        '���ش���
        .Str���� = GetSetting("ZLSOFT", strReg, "����", "1,2,3")
        strWins_temp = "'" & Replace(.Str����, ",", "','") & "'"
        '�Դ��ںŽ�����������
        strSql = "Select TO_CHAR(WMSYS.WM_CONCAT(����)) ���� From (select ���� from ��ҩ���� where ҩ��ID=[1] and ���� in (" & strWins_temp & ") order by ����)"
        Set rsWin = gobjDatabase.OpenSQLRecord(strSql, "", mlngҩ��ID)
        If rsWin.RecordCount > 0 Then
            .Str���� = Nvl(rsWin!����)
        End If
        
        '������Ļ��Ϣ
        .dblLeft = GetSetting("ZLSOFT", strReg, "��", "1024")
        .dblTop = GetSetting("ZLSOFT", strReg, "��", "0")
        .dblWidth = GetSetting("ZLSOFT", strReg, "���", "1024")
        .dblHeight = GetSetting("ZLSOFT", strReg, "�߶�", "768")
        
        '�����е�������ɫ
        .lng������������ɫ = GetSetting("ZLSOFT", strReg, "��������ɫ", vbGreen)
        
        '����ҩ�б������
        .bln��ʾ����ҩ = (Val(GetSetting("ZLSOFT", strReg, "��ʾ����ҩ", "1")) = 1)
        .bln��ʾ��ҩ��� = (Val(GetSetting("ZLSOFT", strReg, "����ҩ���", "0")) = 1)
        .int����ҩ���� = Val(GetSetting("ZLSOFT", strReg, "����ҩ����", "10"))
        .int����ҩ���� = Val(GetSetting("ZLSOFT", strReg, "����ҩ����", "5"))
        .int����ҩ���� = Val(GetSetting("ZLSOFT", strReg, "����ҩ����", "2"))
        .lng����ҩ������ɫ = GetSetting("ZLSOFT", strReg, "����ҩ��ɫ", vbGreen)
        
        '����ҩ�б������
        .bln��ʾ����ҩ = (Val(GetSetting("ZLSOFT", strReg, "��ʾ����ҩ", "1")) = 1)
        .int����ҩ���� = Val(GetSetting("ZLSOFT", strReg, "����ҩ����", "10"))
        .int����ҩ���� = Val(GetSetting("ZLSOFT", strReg, "����ҩ����", "5"))
        .int����ҩ���� = Val(GetSetting("ZLSOFT", strReg, "����ҩ����", "2"))
        .lng����ҩ������ɫ = GetSetting("ZLSOFT", strReg, "����ҩ��ɫ", vbGreen)
        
        '����ҩ�б������
        .bln��ʾ�ѹ��� = (Val(GetSetting("ZLSOFT", strReg, "��ʾ�ѹ���", "1")) = 1)
        .int�ѹ������� = Val(GetSetting("ZLSOFT", strReg, "�ѹ�������", "5"))
        .int�ѹ������� = Val(GetSetting("ZLSOFT", strReg, "�ѹ�������", "5"))
        .int�ѹ������� = Val(GetSetting("ZLSOFT", strReg, "�ѹ�������", "1"))
        .lng�ѹ���������ɫ = GetSetting("ZLSOFT", strReg, "�ѹ�����ɫ", vbGreen)
        
        .intRowPeople = 5
        .intPage = GetSetting("ZLSOFT", strReg, "��ҳʱ��", "5")
        .intRefTime = GetSetting("ZLSOFT", strReg, "ˢ��ʱ��", "10")
'        .intTimeout = GetSetting("ZLSOFT", strReg, "����ʱ��", "10")
        
        .bln��ʾ���� = (Val(GetSetting("ZLSOFT", strReg, "��ʾ����", "1")) = 1)
        .lng����������ɫ = GetSetting("ZLSOFT", strReg, "������ɫ", vbGreen)
        
        .bln��ʾ�������� = (Val(GetSetting("ZLSOFT", strReg, "��ʾ��������", "1")) = 1)
        .lng��������������ɫ = GetSetting("ZLSOFT", strReg, "����������ɫ", vbBlack)
        
        .str��ʾ���� = GetSetting("ZLSOFT", strReg, "��ʾ����", "")
    End With
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************LoadPara*************")
        Call SaveErrLog("������ʱ��:" & gobjDatabase.Currentdate)
        Call SaveErrLog("������Ϣ:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub


Private Sub InitData(ByVal intPage As Integer, ByVal blnRef As Boolean)
'***********************************************************************
'
'ˢ�����ݣ�intPage=1Ϊ�������ʱ���������ݣ�intpage=2Ϊtimer�¼�ˢ������
'
'************************************************************************
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim strpeople As String
    Dim count As Integer
    Dim intcol As Integer
    Dim intSum As Integer
    Dim strTemp As String
    Dim intTemp As Integer
    Dim intCurPage As Integer
    Dim intPraPage As Integer
    Dim intCallPage As Integer
    Dim rsTemp As Recordset
    Dim strSql As String
    '���ƴ���ҩ�б�ı߿�
    Dim strReg As String
    On Error GoTo errHandle
    strReg = "����ģ��\ҩ���Ŷӽк�\Һ������"
    If vfgCallingData.Cols = 0 Then Exit Sub
    If mType_para.bln��ʾ����ҩ Or mType_para.bln��ʾ����ҩ Or mType_para.bln��ʾ�ѹ��� Then
        vfgCallingData.Select mintRows - 1, 0, mintRows - 1, (mintCols) * mRowRec - 1
        'vfgCallingData.CellBorder &HFF00&, -1, -1, -1, 1, 0, 1
        vfgCallingData.CellBorder GetSetting("ZLSOFT", strReg, "������ɫ", vbGreen), -1, -1, -1, 1, 0, 1
    End If
'    strsql = "Select ���� From ���ű� Where ID=[1]"
'    Set rsTemp = gobjDatabase.OpenSQLRecord(strsql, "", mlngҩ��ID)
    For k = 0 To mintCols - 1
        'wwx timerCall ˢ��
        If intPage = 2 Or blnRef Then
            loadCalling (Split(mstrWins, ",")(k))
            For j = 1 To mRowRec
                intSum = intSum + 1
                'Me.vfgCallingData.TextMatrix(0, intSum - 1) = rsTemp!���� & "  " & Split(mstrWins, ",")(k)
                Me.vfgCallingData.Cell(flexcpFontSize, 0, intSum - 1, 0, intSum - 1) = Val(GetSetting("ZLSOFT", strReg, "�ֺ�(1)", "14"))
                Me.vfgCallingData.Cell(flexcpForeColor, 0, intSum - 1, 0, intSum - 1) = GetSetting("ZLSOFT", strReg, "������ɫ", vbGreen)
                Me.vfgCallingData.Cell(flexcpFontName, 0, intSum - 1, 0, intSum - 1) = GetSetting("ZLSOFT", strReg, "����(1)", "����")
                Me.vfgCallingData.Cell(flexcpFontBold, 0, intSum - 1, 0, intSum - 1) = GetSetting("ZLSOFT", strReg, "����(1)", "false")
                Me.vfgCallingData.Cell(flexcpFontItalic, 0, intSum - 1, 0, intSum - 1) = GetSetting("ZLSOFT", strReg, "б��(1)", "false")
                Me.vfgCallingData.TextMatrix(0, intSum - 1) = Split(mstrWins, ",")(k)
                If k Mod 2 = 0 Then
                    strTemp = String(0, " ")
                Else
                    strTemp = String(1, " ")
                End If
                Me.vfgCallingData.Cell(flexcpFontSize, 1, intSum - 1, 1, intSum - 1) = Val(GetSetting("ZLSOFT", strReg, "�ֺ�(0)", "14"))
                Me.vfgCallingData.Cell(flexcpForeColor, 1, intSum - 1, 1, intSum - 1) = GetSetting("ZLSOFT", strReg, "��������ɫ", vbGreen)
                Me.vfgCallingData.Cell(flexcpFontName, 1, intSum - 1, 1, intSum - 1) = GetSetting("ZLSOFT", strReg, "����(0)", "����")
                Me.vfgCallingData.Cell(flexcpFontBold, 1, intSum - 1, 1, intSum - 1) = GetSetting("ZLSOFT", strReg, "����(0)", "false")
                Me.vfgCallingData.Cell(flexcpFontItalic, 1, intSum - 1, 1, intSum - 1) = GetSetting("ZLSOFT", strReg, "б��(0)", "false")
                If Not mrsCallingData.EOF Then
                    SaveDebug (mrsCallingData.RecordCount)
                    SaveDebug ("���ں���������" & mrsCallingData!����)
                    Me.vfgCallingData.TextMatrix(1, intSum - 1) = strTemp & "�� " & mrsCallingData!���� & " ��ҩ"
                Else
                    Me.vfgCallingData.TextMatrix(1, intSum - 1) = strTemp & "�޺�����Ա"
                End If
            Next
        End If
        
        If intPage = 1 Or blnRef Then
            loadCalling (Split(mstrWins, ",")(k))
            For j = 1 To mRowRec
                intSum = intSum + 1
                'Me.vfgCallingData.TextMatrix(0, intSum - 1) = rsTemp!���� & "  " & Split(mstrWins, ",")(k)
                Me.vfgCallingData.Cell(flexcpFontSize, 0, intSum - 1, 0, intSum - 1) = Val(GetSetting("ZLSOFT", strReg, "�ֺ�(1)", "14"))
                Me.vfgCallingData.Cell(flexcpForeColor, 0, intSum - 1, 0, intSum - 1) = GetSetting("ZLSOFT", strReg, "������ɫ", vbGreen)
                Me.vfgCallingData.Cell(flexcpFontName, 0, intSum - 1, 0, intSum - 1) = GetSetting("ZLSOFT", strReg, "����(1)", "����")
                Me.vfgCallingData.Cell(flexcpFontBold, 0, intSum - 1, 0, intSum - 1) = GetSetting("ZLSOFT", strReg, "����(1)", "false")
                Me.vfgCallingData.Cell(flexcpFontItalic, 0, intSum - 1, 0, intSum - 1) = GetSetting("ZLSOFT", strReg, "б��(1)", "false")
                Me.vfgCallingData.TextMatrix(0, intSum - 1) = Split(mstrWins, ",")(k)
                If k Mod 2 = 0 Then
                    strTemp = String(0, " ")
                Else
                    strTemp = String(1, " ")
                End If
                Me.vfgCallingData.Cell(flexcpFontSize, 1, intSum - 1, 1, intSum - 1) = Val(GetSetting("ZLSOFT", strReg, "�ֺ�(0)", "14"))
                Me.vfgCallingData.Cell(flexcpForeColor, 1, intSum - 1, 1, intSum - 1) = GetSetting("ZLSOFT", strReg, "��������ɫ", vbGreen)
                Me.vfgCallingData.Cell(flexcpFontName, 1, intSum - 1, 1, intSum - 1) = GetSetting("ZLSOFT", strReg, "����(0)", "����")
                Me.vfgCallingData.Cell(flexcpFontBold, 1, intSum - 1, 1, intSum - 1) = GetSetting("ZLSOFT", strReg, "����(0)", "false")
                Me.vfgCallingData.Cell(flexcpFontItalic, 1, intSum - 1, 1, intSum - 1) = GetSetting("ZLSOFT", strReg, "б��(0)", "false")
                If Not mrsCallingData.EOF Then
                    SaveDebug (mrsCallingData.RecordCount)
                    SaveDebug ("���ں���������" & mrsCallingData!����)
                    Me.vfgCallingData.TextMatrix(1, intSum - 1) = strTemp & "�� " & mrsCallingData!���� & " ��ҩ"
                Else
                    Me.vfgCallingData.TextMatrix(1, intSum - 1) = strTemp & "�޺�����Ա"
                End If
                Me.vfgCallingData.Cell(flexcpFontSize, 2, intSum - 1, 2, intSum - 1) = Val(GetSetting("ZLSOFT", strReg, "�ֺ�(1)", "14"))
                Me.vfgCallingData.Cell(flexcpForeColor, 2, intSum - 1, 2, intSum - 1) = GetSetting("ZLSOFT", strReg, "������ɫ", vbGreen)
                Me.vfgCallingData.Cell(flexcpFontName, 2, intSum - 1, 2, intSum - 1) = GetSetting("ZLSOFT", strReg, "����(1)", "����")
                Me.vfgCallingData.Cell(flexcpFontBold, 2, intSum - 1, 2, intSum - 1) = GetSetting("ZLSOFT", strReg, "����(1)", "false")
                Me.vfgCallingData.Cell(flexcpFontItalic, 2, intSum - 1, 2, intSum - 1) = GetSetting("ZLSOFT", strReg, "б��(1)", "false")
                If mType_para.bln��ʾ����ҩ Then
                    If intSum - k * mRowRec <= mType_para.int����ҩ���� Then
                        Me.vfgCallingData.TextMatrix(2, intSum - 1) = "����ҩ"
                    End If
                End If
                If mType_para.bln��ʾ����ҩ Then
                    If mType_para.bln��ʾ����ҩ Then
                        If intSum - k * mRowRec <= mType_para.int����ҩ���� + mType_para.int����ҩ���� And intSum - k * mRowRec > mType_para.int����ҩ���� Then
                            Me.vfgCallingData.TextMatrix(2, intSum - 1) = "����ҩ"
                        End If
                    Else
                        If intSum - k * mRowRec <= mType_para.int����ҩ���� Then
                            Me.vfgCallingData.TextMatrix(2, intSum - 1) = "����ҩ"
                        End If
                    End If
                End If
                If mType_para.bln��ʾ�ѹ��� Then
                    If mType_para.bln��ʾ����ҩ And mType_para.bln��ʾ����ҩ Then
                        If intSum - k * mRowRec <= mType_para.int����ҩ���� + mType_para.int����ҩ���� + mType_para.int�ѹ������� And intSum - k * mRowRec > mType_para.int����ҩ���� + mType_para.int����ҩ���� Then
                            Me.vfgCallingData.TextMatrix(2, intSum - 1) = "����"
                        End If
                    Else
                        If Not mType_para.bln��ʾ����ҩ And mType_para.bln��ʾ����ҩ Then
                            If intSum - k * mRowRec <= mType_para.int����ҩ���� + mType_para.int�ѹ������� And intSum - k * mRowRec > mType_para.int����ҩ���� Then
                                Me.vfgCallingData.TextMatrix(2, intSum - 1) = "����"
                            End If
                        Else
                            If mType_para.bln��ʾ����ҩ And Not mType_para.bln��ʾ����ҩ Then
                                If intSum - k * mRowRec <= mType_para.int����ҩ���� + mType_para.int�ѹ������� And intSum - k * mRowRec > mType_para.int����ҩ���� Then
                                    Me.vfgCallingData.TextMatrix(2, intSum - 1) = "����"
                                End If
                            Else
                                If intSum - k * mRowRec <= mType_para.int�ѹ������� Then
                                    Me.vfgCallingData.TextMatrix(2, intSum - 1) = "����"
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If
        If blnRef = False Then
            '��մ���ҩ�б������
            Me.vfgCallingData.Cell(flexcpText, 3, k * mRowRec, mIntCallRow + 2, (k + 1) * mRowRec - 1) = ""
            '��ʾ����ҩ��Ϣ
            ShowPra k, intPage
            '��ʾ����ҩ��Ϣ
            If mType_para.bln��ʾ����ҩ Then
                loadData (Split(mstrWins, ",")(k)), k
                ShowSend k, intPage
            End If
            '��ʾ�ѹ�����Ϣ
            ShowTimeout k, intPage
        End If
        If k <= mintCols - 1 And vfgCallingData.Rows > 2 Then
            '��,��,��,�ײ�,��ֱ,ˮƽ��
            '����ҩ����ҩ������֮��ķָ���
            If mType_para.bln��ʾ����ҩ Then
                vfgCallingData.Select 3, mType_para.int����ҩ���� - 1 + k * mRowRec, mintRows - 1, mType_para.int����ҩ���� - 1 + k * mRowRec
                vfgCallingData.CellBorder GetSetting("ZLSOFT", strReg, "������ɫ", vbGreen), 0, 0, 1, 0, 1, 0
                vfgCallingData.Select mintRows - 1, mType_para.int����ҩ���� - 1 + k * mRowRec, mintRows - 1, mType_para.int����ҩ���� - 1 + k * mRowRec
                vfgCallingData.CellBorder GetSetting("ZLSOFT", strReg, "������ɫ", vbGreen), 0, 0, 1, 1, 1, 1
                If mType_para.bln��ʾ����ҩ And mType_para.bln��ʾ�ѹ��� Then
                    vfgCallingData.Select 3, mType_para.int����ҩ���� + mType_para.int����ҩ���� - 1 + k * mRowRec, mintRows - 1, mType_para.int����ҩ���� + mType_para.int����ҩ���� - 1 + k * mRowRec
                    vfgCallingData.CellBorder GetSetting("ZLSOFT", strReg, "������ɫ", vbGreen), 0, 0, 1, 0, 1, 0
                    vfgCallingData.Select mintRows - 1, mType_para.int����ҩ���� + mType_para.int����ҩ���� - 1 + k * mRowRec, mintRows - 1, mType_para.int����ҩ���� + mType_para.int����ҩ���� - 1 + k * mRowRec
                    vfgCallingData.CellBorder GetSetting("ZLSOFT", strReg, "������ɫ", vbGreen), 0, 0, 1, 1, 1, 1
                End If
            Else
                If mType_para.bln��ʾ����ҩ Then
                    vfgCallingData.Select 3, mType_para.int����ҩ���� - 1 + k * mRowRec, mintRows - 1, mType_para.int����ҩ���� - 1 + k * mRowRec
                    vfgCallingData.CellBorder GetSetting("ZLSOFT", strReg, "������ɫ", vbGreen), 0, 0, 1, 0, 1, 0
                    vfgCallingData.Select mintRows - 1, mType_para.int����ҩ���� - 1 + k * mRowRec, mintRows - 1, mType_para.int����ҩ���� - 1 + k * mRowRec
                    vfgCallingData.CellBorder GetSetting("ZLSOFT", strReg, "������ɫ", vbGreen), 0, 0, 1, 1, 1, 1
                End If
            End If
        End If
        '���߿�
        If k <> mintCols - 1 And vfgCallingData.Rows > 2 Then
            vfgCallingData.Select 3, (k + 1) * mRowRec - 1, mintRows - 1, (k + 1) * mRowRec
            vfgCallingData.CellBorder GetSetting("ZLSOFT", strReg, "������ɫ", vbGreen), -1, -1, -1, 0, 1, 0
           
            If mType_para.bln��ʾ����ҩ Or mType_para.bln��ʾ����ҩ Or mType_para.bln��ʾ�ѹ��� Then
                vfgCallingData.Select mIntCallRow + 2, (k + 1) * mRowRec - 1, mIntCallRow + 2, (k + 1) * mRowRec
                vfgCallingData.CellBorder GetSetting("ZLSOFT", strReg, "������ɫ", vbGreen), -1, -1, -1, 1, 1, 1
            End If
        End If
    Next
    
    '�ϲ����ںͽк���Ϣ
    vfgCallingData.MergeRow(0) = True
    vfgCallingData.MergeRow(1) = True
    vfgCallingData.MergeRow(2) = True
    vfgCallingData.Refresh
    
    vfgCallingData.Select 0, 0, 2, mintCols * mRowRec - 1
    'vfgCallingData.CellBorder &HFF00&, 0, 0, 0, 1, 1, 1
    vfgCallingData.CellBorder GetSetting("ZLSOFT", strReg, "������ɫ", vbGreen), 0, 0, 0, 1, 1, 1
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************InitData*************")
        Call SaveErrLog("������ʱ��:" & gobjDatabase.Currentdate)
        Call SaveErrLog("������Ϣ:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub


Private Sub Form_Load()
    Dim strReg As String
    On Error GoTo errHandle
    '���ز���
    LoadPara
    
    SetFacePostion
    strReg = "����ģ��\ҩ���Ŷӽк�\Һ������"
'    If 25 * Val(GetSetting("ZLSOFT", strReg, "�ֺ�(4)", "14")) > Round(Me.ScaleHeight * 0.9) Then
    Me.vfgCallingData.Move 0, 0, Me.ScaleWidth, IIf(mType_para.bln��ʾ��������, Round(Me.ScaleHeight - 30 * Val(GetSetting("ZLSOFT", strReg, "�ֺ�(4)", "14"))), Round(Me.ScaleHeight))
    Me.BackColor = vbBlack
    Me.lblmsg.Move 0, Me.vfgCallingData.Height + 100, Me.vfgCallingData.Width, Me.ScaleHeight

    '����ģʽȷ���������ʾ���ڣ��������Ǵ��Σ��ര�����ڲ�������ʱ����ѡ��
    If mType_para.bln��������ʾģʽ = False Then
        mstrWins = mType_para.Str����
        If mstrWins = "" Then
            Exit Sub
        End If
    Else
        Me.TimerCall.Enabled = False
    End If
    
    mintCols = UBound(Split(mstrWins, ",")) + 1
'    If mintCols = 0 Then Exit Sub
    mRowRec = mType_para.intRowPeople
    
    'ȷ�����ݼ�����ĳ���
    ReDim mrsData(mintCols)
    ReDim mrsPreparingData(mintCols)
    ReDim mrsTimeoutData(mintCols)
    ReDim mstrSendNames(mintCols)
    ReDim mstrPraNames(mintCols)
    ReDim mstrTimeoutNames(mintCols)
    '��ʼ�����
    InitVSF
    
    InitData 1, False
    
    Me.timerPage.Interval = mType_para.intPage * 1000
    Me.timerLCD.Interval = mType_para.intRefTime * 1000

    Me.lblmsg.Visible = mType_para.bln��ʾ��������
    Me.lblmsg.Caption = IIf(mType_para.str��ʾ���� = "", "ף�����տ�����", mType_para.str��ʾ����) & "   " & Format(gobjDatabase.Currentdate, "yyyy-mm-dd  hh:mm")
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************Load*************")
        Call SaveErrLog("������ʱ��:" & gobjDatabase.Currentdate)
        Call SaveErrLog("������Ϣ:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub

Private Sub loadData(ByVal strWin As String, ByVal Index As Integer)
'************************************************************************
'
'���ش���ҩ�б������
'
'************************************************************************
    Dim strSql As String
    Dim date��ʼ���� As Date
    Dim date�������� As Date
        
    On Error GoTo errHandle
    date��ʼ���� = gobjDatabase.Currentdate
    date��ʼ���� = CDate(Format(date��ʼ����, "yyyy-mm-dd") & " 00:00:00")

    date�������� = gobjDatabase.Currentdate
    date�������� = CDate(Format(date��������, "yyyy-mm-dd") & " 23:59:59")

    strSql = "Select A.����ID,A.����,B.��ҩ����,B.ǩ��ʱ��,B.�������� " & _
             "From δ��ҩƷ��¼ A,ҩƷ�շ���¼ B,��������˵�� D " & _
             "Where A.����=B.���� And A.No=B.NO And A.�ⷿid=B.�ⷿid And (A.����=8 or A.����=9 or A.����=10) and A.�ⷿid=D.����ID " & _
             "  And D.�������� in ('��ҩ��','��ҩ��')"
             '& IIf(mType_para.bln��ʾ�ѹ���, " And round((SYSDATE-Nvl(A.����ʱ��,SYSDATE))*24*60*60)<=" & mType_para.intTimeout, "")
    If mbln��ҩ Then
        strSql = strSql & " and (A.�Ŷ�״̬=2 " & IIf(mType_para.bln��ʾ�ѹ���, "", " or A.�Ŷ�״̬=4") & ") and A.�ⷿid=[1] and A.��ҩ����=[2] and A.�������� between [3] and [4] And (B.��¼״̬=1 Or Mod(B.��¼״̬,3)=0)"
        'strSQL = strSQL & " and (A.�Ŷ�״̬=2 or A.�Ŷ�״̬=4) and A.�ⷿid=52 and A.��ҩ����='����1' and A.�������� >sysdate-2 And (B.��¼״̬=1 Or Mod(B.��¼״̬,3)=0) "
    ElseIf mbln��ҩȷ�� And mbln��ҩ = False Then
        strSql = strSql & " and (A.�Ŷ�״̬=1 or A.�Ŷ�״̬=2 " & IIf(mType_para.bln��ʾ�ѹ���, "", " or A.�Ŷ�״̬=4") & ") and A.�ⷿid=[1] and A.��ҩ����=[2] and A.�������� between [3] and [4] And (B.��¼״̬=1 Or Mod(B.��¼״̬,3)=0)"
        'strSQL = strSQL & " and (A.�Ŷ�״̬=1 or A.�Ŷ�״̬=2 or A.�Ŷ�״̬=4) and A.�ⷿid=52 and A.��ҩ����='����1' and A.�������� >sysdate-2And (B.��¼״̬=1 Or Mod(B.��¼״̬,3)=0) "
    ElseIf mbln��ҩ = False And mbln��ҩȷ�� = False Then
        strSql = strSql & "  and (A.�Ŷ�״̬<3 or A.�Ŷ�״̬ is null " & IIf(mType_para.bln��ʾ�ѹ���, "", " or A.�Ŷ�״̬=4") & ") and A.�ⷿid=[1] and A.��ҩ����=[2] and A.�������� between [3] and [4] And (B.��¼״̬=1 Or Mod(B.��¼״̬,3)=0)"
        'strSQL = strSQL & "  and (A.�Ŷ�״̬<>3 or A.�Ŷ�״̬ is null) and A.�ⷿid=52 and A.��ҩ����='����1' and A.��������>sysdate-2 And (B.��¼״̬=1 Or Mod(B.��¼״̬,3)=0) "
    End If
    strSql = "Select Rownum ���,����,���� " & _
             "From ( " & _
                    "Select min(" & IIf(mbln��ҩ, "��ҩ����", "Nvl(ǩ��ʱ��,��������)") & ") ����,����id,���� " & _
                    "From (" & strSql & ") " & _
                    "Where ����ID Not In (Select distinct A.����ID From δ��ҩƷ��¼ A,ҩƷ�շ���¼ B,������ü�¼ C " & _
                                         "Where A.����=B.���� And A.No=B.NO And A.�ⷿid=B.�ⷿid and B.����id=C.id and (A.����=8 or A.����=9 or A.����=10) " & _
                                         "  and A.�Ŷ�״̬=4 and A.�ⷿid=[1] and A.��ҩ����=[2] and A.�������� between [3] and [4] And (B.��¼״̬=1 Or Mod(B.��¼״̬,3)=0)) " & _
                    "Group By ����,����id " & _
                    "Order by ���� " & _
                    ")"
    'Call SaveErrLog(strSQL)
    Set mrsData(Index) = gobjDatabase.OpenSQLRecord(strSql, "", mlngҩ��ID, strWin, date��ʼ����, date��������)
    'Call SaveErrLog(mlngҩ��ID & "," & strWin & "," & date��ʼ���� & "," & date��������)
    'If mrsData(Index).State = 1 Then mrsData(Index).Close
    'mrsData(Index).Open strSQL, gcnOracle
'    If Not mrsData(Index).EOF Then
'        If Nvl(mrsData(Index)!��ҩ����) <> "" Then
'            mrsData(Index).Sort = "��ҩ����"
'        Else
'            mrsData(Index).Sort = "ǩ��ʱ��"
'        End If
'    End If
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************LoadData*************")
        Call SaveErrLog("������ʱ��:" & gobjDatabase.Currentdate)
        Call SaveErrLog("������Ϣ:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub

Private Sub loadCalling(ByVal strWin As String)
'************************************************************************
'
'���ص�ǰ���е�����
'
'************************************************************************
    Dim strSql As String
    Dim date��ʼ���� As Date
    Dim date�������� As Date
    
    On Error GoTo errHandle
    date��ʼ���� = gobjDatabase.Currentdate
    'date��ʼ���� = Now - 1
    date��ʼ���� = CDate(Format(date��ʼ����, "yyyy-mm-dd") & " 00:00:00")
    
    date�������� = gobjDatabase.Currentdate
    'date�������� = Now
    date�������� = CDate(Format(date��������, "yyyy-mm-dd") & " 23:59:59")
    
    strSql = "select ���� from δ��ҩƷ��¼ where �Ŷ�״̬=3 and �ⷿid=[1] and ��ҩ����=[2] and �������� between [3] and [4]"
    Set mrsCallingData = gobjDatabase.OpenSQLRecord(strSql, "", mlngҩ��ID, strWin, date��ʼ����, date��������)
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************LoadCalling*************")
        Call SaveErrLog("������ʱ��:" & gobjDatabase.Currentdate)
        Call SaveErrLog("������Ϣ:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub


Private Sub loadPreparing(ByVal strWin As String, ByVal intIndex As Integer)
'************************************************************************
'
'���ش���ҩ�б������
'
'************************************************************************
    Dim strSql As String
    
    Dim date��ʼ���� As Date
    Dim date�������� As Date
        
    On Error GoTo errHandle
    date��ʼ���� = gobjDatabase.Currentdate
    date��ʼ���� = CDate(Format(date��ʼ����, "yyyy-mm-dd") & " 00:00:00")

    date�������� = gobjDatabase.Currentdate
    date�������� = CDate(Format(date��������, "yyyy-mm-dd") & " 23:59:59")
    
    strSql = "Select rownum ���,A.����,B.��������,B.ǩ��ʱ�� From δ��ҩƷ��¼ A,ҩƷ�շ���¼ B,������ü�¼ C" & _
             " Where A.����=B.���� And A.No=B.NO And A.�ⷿid=B.�ⷿid and B.����id=C.id and (A.����=8 or A.����=9 or A.����=10) "
    If mbln��ҩȷ�� Then
        strSql = strSql & "and A.�Ŷ�״̬=1 and A.�ⷿid=[1] and A.��ҩ����=[2] and A.�������� between [3] and [4] And (B.��¼״̬=1 Or Mod(B.��¼״̬,3)=0)"
    Else
        strSql = strSql & "and (A.�Ŷ�״̬=1 or A.�Ŷ�״̬=0 or A.�Ŷ�״̬ is null) and A.�ⷿid=[1] and A.��ҩ����=[2] and A.�������� between [3] and [4] And (B.��¼״̬=1 Or Mod(B.��¼״̬,3)=0)"
    End If
    'strSQL = strSQL & "and (A.�Ŷ�״̬=1 or A.�Ŷ�״̬=0 or A.�Ŷ�״̬ is null) and A.�ⷿid=52 and A.��ҩ����='����1' and A.��������>sysdate-2 And (B.��¼״̬=1 Or Mod(B.��¼״̬,3)=0)"
    If mbln��ҩ = False Then strSql = strSql & " And 1=2"
    strSql = "Select ����,min(���) ��� From (" & strSql & " Order by Nvl(B.ǩ��ʱ��,A.��������)) group by ���� order by ���"
    Set mrsPreparingData(intIndex) = gobjDatabase.OpenSQLRecord(strSql, "", mlngҩ��ID, strWin, date��ʼ����, date��������)
    
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************LoadPreparing*************")
        Call SaveErrLog("������ʱ��:" & gobjDatabase.Currentdate)
        Call SaveErrLog("������Ϣ:" & err.Number & "->" & err.Description)
        Call SaveErrLog("strSQL:" & strSql)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub
Private Sub loadTimeout(ByVal strWin As String, ByVal intIndex As Integer)
'************************************************************************
'
'�����ѹ��ţ�����ҩ���������룩�б������
'
'************************************************************************
    Dim strSql As String
    
    Dim date��ʼ���� As Date
    Dim date�������� As Date
        
    On Error GoTo errHandle
    date��ʼ���� = gobjDatabase.Currentdate
    date��ʼ���� = CDate(Format(date��ʼ����, "yyyy-mm-dd") & " 00:00:00")

    date�������� = gobjDatabase.Currentdate
    date�������� = CDate(Format(date��������, "yyyy-mm-dd") & " 23:59:59")
    
    strSql = "Select distinct A.����ID,A.����,A.����ʱ�� From δ��ҩƷ��¼ A,ҩƷ�շ���¼ B,������ü�¼ C" & _
             " Where A.����=B.���� And A.No=B.NO And A.�ⷿid=B.�ⷿid and B.����id=C.id and (A.����=8 or A.����=9 or A.����=10) " & _
             "   and A.�Ŷ�״̬=4 and A.�ⷿid=[1] and A.��ҩ����=[2] and A.�������� between [3] and [4] And (B.��¼״̬=1 Or Mod(B.��¼״̬,3)=0) " & _
             " union all " & _
             " Select distinct A.����ID,A.����,A.����ʱ�� From δ��ҩƷ��¼ A,ҩƷ�շ���¼ B,סԺ���ü�¼ C" & _
             " Where A.����=B.���� And A.No=B.NO And A.�ⷿid=B.�ⷿid and B.����id=C.id and (A.����=8 or A.����=9 or A.����=10) " & _
             "   and A.�Ŷ�״̬=4 and A.�ⷿid=[1] and A.��ҩ����=[2] and A.�������� between [3] and [4] And (B.��¼״̬=1 Or Mod(B.��¼״̬,3)=0)"
'    If mbln��ҩ = False Then
'        strSQL = strSQL & " and round((sysdate-Nvl(B.ǩ��ʱ��,B.��������))*24*60)>" & mType_para.intTimeout
'        strSQL = strSQL & " Order by Nvl(B.ǩ��ʱ��,B.��������) DESC"
'    Else
'        strSQL = strSQL & " and round((sysdate-B.��ҩ����)*24*60)>" & mType_para.intTimeout
'        strSQL = strSQL & " Order by B.��ҩ���� DESC"
'    End If
'    strSQL = strSQL & " And A.����ʱ�� IS NOT NULL And round((SYSDATE-Nvl(A.����ʱ��,SYSDATE))*24*60*60)>" & mType_para.intTimeout
    strSql = "Select rownum ���,���� " & _
             "From (Select ����ID,����,min(����ʱ��) ����ʱ�� " & _
                   "From (" & strSql & ") " & _
                   "Group by ����ID,���� " & _
                   "Order by ����ʱ�� asc " & _
                   ")"
    'Call SaveErrLog(strSQL)
    Set mrsTimeoutData(intIndex) = gobjDatabase.OpenSQLRecord(strSql, "", mlngҩ��ID, strWin, date��ʼ����, date��������)
    'Call SaveErrLog(mlngҩ��ID & "," & strWin & "," & date��ʼ���� & "," & date��������)
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************LoadTimeout*************")
        Call SaveErrLog("������ʱ��:" & gobjDatabase.Currentdate)
        Call SaveErrLog("������Ϣ:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub
Private Sub InitVSF()
'************************************************************************
'
'��ʼ�����
'
'************************************************************************
    Dim intColWidth As Integer
    Dim lngRowheight As Long, lngRowheights As Long
    Dim i As Integer
    Dim strReg As String
    Dim dblHeight As Double
    On Error GoTo errHandle
    strReg = "����ģ��\ҩ���Ŷӽк�\Һ������"
    
    mintRows = 3

    
    mIntCallRow = IIf(mType_para.bln��ʾ����ҩ, mType_para.int����ҩ����, 0)
    mIntPraRow = IIf(mType_para.bln��ʾ����ҩ, mType_para.int����ҩ����, 0)
    
    mIntCallCol = IIf(mType_para.bln��ʾ����ҩ, mType_para.int����ҩ����, 0)
    mIntPraCol = IIf(mType_para.bln��ʾ����ҩ, mType_para.int����ҩ����, 0)
    mIntTimeoutCol = IIf(mType_para.bln��ʾ�ѹ���, mType_para.int�ѹ�������, 0)
    mRowRec = mIntCallCol + mIntPraCol + mIntTimeoutCol
    'mRowRec = mType_para.int����ҩ����
    'mintRows = mintRows + mIntCallRow + mIntPraRow + IIf(mType_para.bln��ʾ����ҩ, 1, 0) + IIf(mType_para.bln��ʾ����ҩ, 1, 0)
    mintRows = mintRows + mIntCallRow
    
    With vfgCallingData
        .Rows = mintRows
        .Cols = mintCols * mRowRec
        
        If .Cols = 0 Then Exit Sub
        
'        If mintCols = 0 Then
'            Unload Me
'            Exit Sub
'        End If
        '���ñ��Ϊ���ɺϲ�
        .MergeCells = flexMergeFree

         '���������������ɫ��С
        SetFont
        
        intColWidth = Me.ScaleWidth / (mintCols * mRowRec)
        '�������ݾ�����ʾ
        For i = 0 To mintCols * mRowRec - 1
            .ColWidth(i) = intColWidth
            vfgCallingData.ColAlignment(i) = flexAlignCenterCenter
        Next
        
        If mType_para.bln��ʾ����ҩ Or mType_para.bln��ʾ����ҩ Or mType_para.bln��ʾ�ѹ��� Then
            .RowHeight(0) = 25 * Val(GetSetting("ZLSOFT", strReg, "�ֺ�(1)", "14"))
            .RowHeight(1) = 30 * Val(GetSetting("ZLSOFT", strReg, "�ֺ�(0)", "14"))
            .RowHeight(2) = 25 * Val(GetSetting("ZLSOFT", strReg, "�ֺ�(1)", "14"))
        Else
            .RowHeight(0) = 25 * Val(GetSetting("ZLSOFT", strReg, "�ֺ�(1)", "14"))
            .RowHeight(1) = IIf(mType_para.bln��ʾ����, .Height - .RowHeight(0), .Height)
            .RowHeight(2) = 0
        End If
        
        If Not mType_para.bln��ʾ���� Then
            .RowHeight(0) = 0
        End If
        If vfgCallingData.Rows > 3 Then
            'mIntCallRow����ҩ����,�ܵ�����Ӧ����mIntCallRow����ҩ����+3
            lngRowheight = Round((.Height - .RowHeight(0) - .RowHeight(1) - .RowHeight(2)) / mIntCallRow)
            lngRowheights = 0
            For i = 3 To .Rows - 1
                .RowHeight(i) = lngRowheight
                lngRowheights = lngRowheights + lngRowheight
            Next
            vfgCallingData.Height = lngRowheights + .RowHeight(0) + .RowHeight(1) + .RowHeight(2) + 50
        End If
    End With
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************InitVSF*************")
        Call SaveErrLog("������ʱ��:" & gobjDatabase.Currentdate)
        Call SaveErrLog("������Ϣ:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub

Private Sub Form_Resize()
    Dim strReg As String
    On Error GoTo errHandle
    
    strReg = "����ģ��\ҩ���Ŷӽк�\Һ������"
    'Me.vfgCallingData.Move 0, 0, Me.ScaleWidth, IIf(mType_para.bln��ʾ��������, Round(Me.ScaleHeight * 0.9), Round(Me.ScaleHeight))
    'Me.PicMsg.Move 0, Me.vfgCallingData.Height, Me.vfgCallingData.Width, Round(Me.ScaleHeight * 0.1)
    Me.vfgCallingData.Move 0, 0, Me.ScaleWidth, IIf(mType_para.bln��ʾ��������, Round(Me.ScaleHeight - 30 * Val(GetSetting("ZLSOFT", strReg, "�ֺ�(4)", "14"))), Round(Me.ScaleHeight))
    Me.lblmsg.Move 0, Me.vfgCallingData.Height + 100, Me.vfgCallingData.Width, Me.ScaleHeight
    'Me.lblmsg.Move 0, Me.PicMsg.Height / 20, Me.PicMsg.Width, Me.PicMsg.Height
    InitVSF
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************Form_Resize*************")
        Call SaveErrLog("������ʱ��:" & gobjDatabase.Currentdate)
        Call SaveErrLog("������Ϣ:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    On Error GoTo errHandle
    For i = 0 To mintCols - 1
'        mstrSendNames(i) = ""
'        mstrPraNames(i) = ""
        Set mrsData(i) = Nothing
        Set mrsPreparingData(i) = Nothing
        Set mrsTimeoutData(i) = Nothing
    Next
    Set mrsCallingData = Nothing
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************Form_Unload*************")
        Call SaveErrLog("������ʱ��:" & gobjDatabase.Currentdate)
        Call SaveErrLog("������Ϣ:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub

Private Sub TimerCall_Timer()
    InitData 2, True
End Sub

Private Sub timerPage_Timer()
'************************************************************************
'
'�Դ���ҩ�б�����ݽ��з�ҳ
'
'************************************************************************
    Dim i As Integer
    Dim intcol As Integer
    Dim k As Integer
    Dim count As Integer
    Dim intPage As Integer
    Dim strTemp As String
    Dim intCallPage As Integer
    Dim intPraPage As Integer
    Dim strReg As String
    On Error GoTo errHandle
    strReg = "����ģ��\ҩ���Ŷӽк�\Һ������"
    If mType_para.bln��ʾ����ҩ = False And mType_para.bln��ʾ����ҩ = False And mType_para.bln��ʾ�ѹ��� = False Then Exit Sub

'    Me.timerLCD.Enabled = False

    For k = 0 To mintCols - 1
        intcol = k * mRowRec
        '����������ڷ�ҳ֮���ҳ��
'        For intcol = k * mRowRec To (k + 1) * mRowRec - 1
        If (intcol \ mRowRec) Mod 2 = 0 Then
            strTemp = String(0, " ")
        Else
            strTemp = String(1, " ")
        End If

'        If mType_para.bln��ʾ����ҩ = True Then
'            '����ҩ����ҳ��
'            'intCallPage = (mrsData(k).RecordCount \ (mRowRec * mIntCallRow) + IIf(mrsData(k).RecordCount Mod (mRowRec * mIntCallRow) = 0, 0, 1))
'            intCallPage = 1
'            If intCallPage = 0 Then intCallPage = 1
'
'            '��ǰҳ
'            'intPage = Val(Mid(Me.vfgCallingData.TextMatrix(mIntCallRow + 2, intcol), 4, InStr(1, Me.vfgCallingData.TextMatrix(mIntCallRow + 2, intcol), "/"))) + 1
'            intPage = 1
'            If intPage > intCallPage Then intPage = 1
'
'            Me.vfgCallingData.Cell(flexcpText, mIntCallRow + 2, intcol, mIntCallRow + 2, (k + 1) * mRowRec - 1) = "����ҩ   " & strTemp & intPage & "/" & intCallPage & " ��" & mrsData(k).RecordCount & "��"
'        End If

    
        If mType_para.bln��ʾ����ҩ = True Then
            '�����ݼ������ݣ������Ѿ��Ƶ����һ�������˾��Ƶ���һ��
            If mrsData(k).EOF And mrsData(k).RecordCount > 0 Then mrsData(k).MoveFirst
            '��մ���ҩ�б������
            If mType_para.bln��ʾ����ҩ = True Then
                Me.vfgCallingData.Cell(flexcpText, 3, k * mRowRec + mType_para.int����ҩ����, mIntCallRow + 2, k * mRowRec + mType_para.int����ҩ���� + mType_para.int����ҩ���� - 1) = ""
                intcol = k * mRowRec + mType_para.int����ҩ����
            Else
                Me.vfgCallingData.Cell(flexcpText, 3, k * mRowRec, mIntCallRow + 2, k * mRowRec + mType_para.int����ҩ���� - 1) = ""
                intcol = k * mRowRec
            End If

            i = 3
            count = 0
            Do While Not mrsData(k).EOF
                If mrsData(k)!���� <> "" Then
                    count = count + 1
                    Me.vfgCallingData.TextMatrix(i, intcol) = Nvl(mrsData(k)!����)
                    
                    Me.vfgCallingData.Cell(flexcpFontSize, i, intcol, i, intcol) = Val(GetSetting("ZLSOFT", strReg, "�ֺ�(3)", "14"))
                    Me.vfgCallingData.Cell(flexcpForeColor, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "����ҩ��ɫ", vbGreen)
                    Me.vfgCallingData.Cell(flexcpFontName, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "����(3)", "����")
                    Me.vfgCallingData.Cell(flexcpFontBold, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "����(3)", "false")
                    Me.vfgCallingData.Cell(flexcpFontItalic, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "б��(3)", "false")
                    
    '                'ÿ����ʾ���ƶ�������������һ��
                    If intcol < IIf(mType_para.bln��ʾ����ҩ, mType_para.int����ҩ����, 0) + mType_para.int����ҩ���� + k * mRowRec - 1 Then
                        intcol = intcol + 1
                    Else
                        intcol = 0
                        i = i + 1
                    End If
                End If

                mrsData(k).MoveNext
                
                 '�����ݵ���ʾ�Ѿ���������ֵʱ���˳�ѭ��
                If count >= mType_para.int����ҩ���� Then
                    Exit Do
                End If
            Loop
        End If
        
'        intcol = 0
'        If mType_para.bln��ʾ����ҩ = True Then
'             '��ǰ��ҳ��
'            intPraPage = (mrsPreparingData(k).RecordCount \ (mIntPraRow * mRowRec) + IIf(mrsPreparingData(k).RecordCount Mod (mIntPraRow * mRowRec) = 0, 0, 1))
'
'            If intPraPage = 0 Then intPraPage = 1
'
'            intPage = Val(Mid(Me.vfgCallingData.TextMatrix(mintRows - 1, intcol), 4, InStr(1, Me.vfgCallingData.TextMatrix(mintRows - 1, intcol), "/"))) + 1
'
'            If intPage > intPraPage Then intPage = 1
'
'            Me.vfgCallingData.Cell(flexcpText, mintRows - 1, intcol, mintRows - 1, (k + 1) * mRowRec - 1) = "����ҩ   " & strTemp & intPage & "/" & intPraPage & " ��" & mrsPreparingData(k).RecordCount & "��"
'        End If

        If mType_para.bln��ʾ����ҩ = True Then
            count = 0
            '�����ݼ������ݣ������Ѿ��Ƶ����һ�������˾��Ƶ���һ��
            If mrsPreparingData(k).EOF And mrsPreparingData(k).RecordCount > 0 Then mrsPreparingData(k).MoveFirst
            '��մ���ҩ�б������
            Me.vfgCallingData.Cell(flexcpText, 3, k * mRowRec, mIntCallRow + 2, k * mRowRec + mType_para.int����ҩ���� - 1) = ""

            i = 3
            intcol = k * mRowRec
            Do While Not mrsPreparingData(k).EOF
                If mrsPreparingData(k)!���� <> "" Then
                    count = count + 1
                    Me.vfgCallingData.TextMatrix(i, intcol) = mrsPreparingData(k)!����
                    Me.vfgCallingData.Cell(flexcpFontSize, i, intcol, i, intcol) = Val(GetSetting("ZLSOFT", strReg, "�ֺ�(2)", "14"))
                    Me.vfgCallingData.Cell(flexcpForeColor, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "����ҩ��ɫ", vbGreen)
                    Me.vfgCallingData.Cell(flexcpFontName, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "����(2)", "����")
                    Me.vfgCallingData.Cell(flexcpFontBold, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "����(2)", "false")
                    Me.vfgCallingData.Cell(flexcpFontItalic, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "б��(2)", "false")

                    'һ�����ݼ�����֮��������һ��
                    If intcol < mType_para.int����ҩ���� + k * mRowRec - 1 Then
                        intcol = intcol + 1
                    Else
                        intcol = 0
                        i = i + 1
                    End If

'                    If count = mRowRec Then
'                        count = 0
'                        intcol = k * mRowRec
'                        If Not mrsPreparingData(k).EOF Then i = i + 1
'                    End If
                End If
                
                mrsPreparingData(k).MoveNext
                '�����ݵ���ʾ�Ѿ���������ֵʱ���˳�ѭ��
                If count >= mType_para.int����ҩ���� Then
                    Exit Do
                End If
            Loop
        End If

        If mType_para.bln��ʾ�ѹ��� = True Then
            '�����ݼ������ݣ������Ѿ��Ƶ����һ�������˾��Ƶ���һ��
            If mrsData(k).EOF And mrsData(k).RecordCount > 0 Then mrsData(k).MoveFirst
            '��մ���ҩ�б������
            Me.vfgCallingData.Cell(flexcpText, 3, k * mRowRec + IIf(mType_para.bln��ʾ����ҩ, mType_para.int����ҩ����, 0) + IIf(mType_para.bln��ʾ����ҩ, mType_para.int����ҩ����, 0), mIntCallRow + 2, k * mRowRec + IIf(mType_para.bln��ʾ����ҩ, mType_para.int����ҩ����, 0) + IIf(mType_para.bln��ʾ����ҩ, mType_para.int����ҩ����, 0) + mType_para.int�ѹ������� - 1) = ""
            intcol = k * mRowRec + IIf(mType_para.bln��ʾ����ҩ, mType_para.int����ҩ����, 0) + IIf(mType_para.bln��ʾ����ҩ, mType_para.int����ҩ����, 0)

            i = 3
            count = 0
            Do While Not mrsData(k).EOF
                If mrsData(k)!���� <> "" Then
                    count = count + 1
                    Me.vfgCallingData.TextMatrix(i, intcol) = Nvl(mrsData(k)!����)
                    
                    Me.vfgCallingData.Cell(flexcpFontSize, i, intcol, i, intcol) = Val(GetSetting("ZLSOFT", strReg, "�ֺ�(5)", "14"))
                    Me.vfgCallingData.Cell(flexcpForeColor, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "�ѹ�����ɫ", vbGreen)
                    Me.vfgCallingData.Cell(flexcpFontName, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "����(5)", "����")
                    Me.vfgCallingData.Cell(flexcpFontBold, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "����(5)", "false")
                    Me.vfgCallingData.Cell(flexcpFontItalic, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "б��(5)", "false")
                    
    '                'ÿ����ʾ���ƶ�������������һ��
                    If intcol < IIf(mType_para.bln��ʾ����ҩ, mType_para.int����ҩ����, 0) + IIf(mType_para.bln��ʾ����ҩ, mType_para.int����ҩ����, 0) + mType_para.int�ѹ������� + k * mRowRec - 1 Then
                        intcol = intcol + 1
                    Else
                        intcol = 0
                        i = i + 1
                    End If
                End If

                mrsData(k).MoveNext
                
                 '�����ݵ���ʾ�Ѿ���������ֵʱ���˳�ѭ��
                If count >= mType_para.int�ѹ������� Then
                    Exit Do
                End If
            Loop
        End If
    Next
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************TimerPage_Timer*************")
        Call SaveErrLog("������ʱ��:" & gobjDatabase.Currentdate)
        Call SaveErrLog("������Ϣ:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub

Public Sub ShowMe(ByVal lngҩ��ID As Long, ByVal strWins As String, ByVal bln��ҩ As Boolean, ByVal bln��ҩȷ�� As Boolean)
'**************************************************************************
'�򿪴���Ľӿڣ�lngҩ��ID����ǰ��ҩ��id��strWins���������Ӵ�
'**************************************************************************
    Dim rsWin As ADODB.Recordset
    Dim strWins_temp As String, strSql As String
    Dim strTemp As String
    Dim strReg As String
    Dim cls As New clsLCDShow
    
    On Error GoTo errHandle
    mlngҩ��ID = lngҩ��ID
    strWins_temp = "'" & Replace(strWins, ",", "','") & "'"
    mstrWins = strWins
    '�Դ��ںŽ�����������
    strSql = "Select TO_CHAR(WMSYS.WM_CONCAT(����)) ���� From (select ���� from ��ҩ���� where ҩ��ID=[1] and ���� in (" & strWins_temp & ") order by ����)"
    Set rsWin = gobjDatabase.OpenSQLRecord(strSql, "", mlngҩ��ID)
    If rsWin.RecordCount > 0 Then
        mstrWins = Nvl(rsWin!����)
    End If
    mbln��ҩ = bln��ҩ
    mbln��ҩȷ�� = bln��ҩȷ��
    
    strReg = "����ģ��\ҩ���Ŷӽк�\Һ������"
    strTemp = GetSetting("ZLSOFT", strReg, "����", "1,2,3")
    If strTemp = "" And strWins = "" Then
        Call SaveErrLog("************************************")
        Call SaveErrLog("������ʱ��:" & gobjDatabase.Currentdate)
        Call SaveErrLog("strTemp:" & strTemp)
        Call SaveErrLog("strWins:" & strWins)
        Call SaveErrLog("��ҩ����Ϊ��,����ʾ�Ŷ���Ļ")
        Call SaveErrLog("************************************")
        cls.zlClose
        Exit Sub
    End If
    Me.Show
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************ShoeMe*************")
        Call SaveErrLog("������ʱ��:" & gobjDatabase.Currentdate)
        Call SaveErrLog("������Ϣ:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub

Private Sub timerLCD_Timer()
'************************************************************************
'
'ˢ�´���ҩ�б������
'
'************************************************************************
    InitData 2, False
'    Dim i As Long, j As Long, lngRowheight As Long, lngRowHeights As Long
'    j = 0
'    For i = 0 To vfgCallingData.Rows - 1
'        j = j + vfgCallingData.RowHeight(i)
'    Next
'    With vfgCallingData
'    lngRowheight = Round((.Height - .RowHeight(0) - .RowHeight(1) - .RowHeight(2)) / mIntCallRow)
'    lngRowHeights = 0
'    For i = 3 To .Rows - 1
'        .Cell(flexcpFontSize, i, 0, i, 14) = 7.5
'        .Cell(flexcpForeColor, i, 0, i, 14) = vbGreen
'        .Cell(flexcpFontName, i, 0, i, 14) = "����"
'        .Cell(flexcpFontBold, i, 0, i, 14) = "false"
'        .Cell(flexcpFontItalic, i, 0, i, 14) = "false"
'        .TextMatrix(i, 0) = .RowHeight(0)
'        .TextMatrix(i, 1) = .RowHeight(1)
'        .TextMatrix(i, 2) = .RowHeight(i)
'        .TextMatrix(i, 3) = .Height
'        .TextMatrix(i, 4) = lblmsg.Height
'        .TextMatrix(i, 5) = Me.ScaleHeight
'        .TextMatrix(i, 6) = Me.Height
'        .TextMatrix(i, 7) = lblmsg.Visible
'        lngRowHeights = lngRowHeights + .RowHeight(i)
'    Next
'        .Height = lngRowHeights + .RowHeight(0) + .RowHeight(1) + .RowHeight(2) + 100
'    End With
    Me.lblmsg.Caption = IIf(mType_para.str��ʾ���� = "", "ף�����տ�����", mType_para.str��ʾ����) & "   " & Format(gobjDatabase.Currentdate, "yyyy-mm-dd  hh:mm")
End Sub

Public Sub ChangeCall(ByVal strWin As String, ByVal strName As String)
'****************************************************************************
'
'���µ�ǰ������Ϣ
'
'**************************************************************************

    InitData 2, True
End Sub

Private Sub ShowSend(ByVal Index As Integer, ByVal intPage As Integer)
'******************************************************************************
'
'������ҩ�����ݼ��ص���������
'
'******************************************************************************
    Dim count As Integer
    Dim i As Integer
    Dim intcol As Integer
    Dim intCallPage As Integer
    Dim intCurPage As Integer
    Dim strTemp As String
    Dim strNames As String
    Dim strReg As String
    On Error GoTo errHandle
    strReg = "����ģ��\ҩ���Ŷӽк�\Һ������"
    '��ʾ����ҩ��Ϣ
    If mType_para.bln��ʾ����ҩ Then
        '�����ܵ�ҳ��
        'intCallPage = (mrsData(Index).RecordCount \ (mRowRec * mIntCallRow) + IIf(mrsData(Index).RecordCount Mod (mRowRec * mIntCallRow) = 0, 0, 1))
        intCallPage = 1
        If intCallPage = 0 Then intCallPage = 1
        
        '�ж��Ƿ�Ϊ�������
        If intPage <> 1 Then
            For i = 0 To Me.vfgCallingData.Cols - 1
                '���㵱ǰҳ��
                If vfgCallingData.TextMatrix(0, i) = (Split(mstrWins, ",")(Index)) Then
                    'intCurPage = Val(Mid(Me.vfgCallingData.TextMatrix(mIntCallRow + 2, i), 4, InStr(1, Me.vfgCallingData.TextMatrix(mIntCallRow + 2, i), "/")))
                    intCurPage = 1
                End If
            Next
            '����¼�����α�����ǰҳ��ʾ������
'            For i = 1 To mIntCallRow * mRowRec * (intCurPage - 1)
'                If Not mrsData(Index).EOF Then
'                    mrsData(Index).MoveNext
'                End If
'            Next
        Else
            intCurPage = 1
        End If
        
        count = 0
        i = 3
        intcol = Index * mRowRec + IIf(mType_para.bln��ʾ����ҩ, mType_para.int����ҩ����, 0)
        
        'ѭ����¼������ʾ����
        Do While Not mrsData(Index).EOF
            If mrsData(Index)!���� <> "" Then
                count = count + 1
                '������ʾ��ָ���������˳�ѭ��
                If count > mType_para.int����ҩ���� Then
                    Exit Do
                End If
                Me.vfgCallingData.Cell(flexcpFontSize, i, intcol, i, intcol) = Val(GetSetting("ZLSOFT", strReg, "�ֺ�(2)", "14"))
                Me.vfgCallingData.Cell(flexcpForeColor, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "����ҩ��ɫ", vbGreen)
                Me.vfgCallingData.Cell(flexcpFontName, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "����(2)", "����")
                Me.vfgCallingData.Cell(flexcpFontBold, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "����(2)", "false")
                Me.vfgCallingData.Cell(flexcpFontItalic, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "б��(2)", "false")
                Me.vfgCallingData.TextMatrix(i, intcol) = TranRowNum(Val(Nvl(mrsData(Index)!���))) & Nvl(mrsData(Index)!����)
                
'                'ÿ����ʾ��ָ��������������һ��
                If intcol < IIf(mType_para.bln��ʾ����ҩ, mType_para.int����ҩ����, 0) + mType_para.int����ҩ���� + Index * mRowRec - 1 Then
                    intcol = intcol + 1
                Else
                    intcol = Index * mRowRec + IIf(mType_para.bln��ʾ����ҩ, mType_para.int����ҩ����, 0)
                    i = i + 1
                End If
'                If count = mRowRec Then
'                    count = 0
'                    intcol = Index * mRowRec
'                    If Not mrsData(Index).EOF Then i = i + 1
'                End If
            End If
            
            'mstrSendNames(Index) = mstrSendNames(Index) & Nvl(mrsData(Index)!����) & ","
            '������һ����¼
            mrsData(Index).MoveNext
'            intcol = intcol + 1
'            If count = mRowRec Then
'                count = 0
'                intcol = Index * mRowRec
'                If Not mrsData(Index).EOF Then i = i + 1
'            End If
            
        Loop
        
    End If
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************ShowSend*************")
        Call SaveErrLog("������ʱ��:" & gobjDatabase.Currentdate)
        Call SaveErrLog("������Ϣ:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub

Private Sub ShowPra(ByVal Index As Integer, ByVal intPage As Integer)
    Dim count As Integer
    Dim i As Integer
    Dim intPraPage As Integer
    Dim intCurPage As Integer
    Dim intcol As Integer
    Dim strTemp As String
    Dim strReg As String
    On Error GoTo errHandle
    strReg = "����ģ��\ҩ���Ŷӽк�\Һ������"
    If mType_para.bln��ʾ����ҩ Then
        '���ش���ҩ��Ϣ
        loadPreparing (Split(mstrWins, ",")(Index)), Index
        
        '������ҳ��
        'intPraPage = (mrsPreparingData(Index).RecordCount \ (mIntPraRow * mRowRec) + IIf(mrsPreparingData(Index).RecordCount Mod (mIntPraRow * mRowRec) = 0, 0, 1))
        intPraPage = 1
        If intPraPage = 0 Then intPraPage = 1
    
        '�ж��Ƿ��Ǵ������
        If intPage <> 1 Then
            For i = 0 To Me.vfgCallingData.Cols - 1
                '�õ���ǰҳ��
                If vfgCallingData.TextMatrix(0, i) = (Split(mstrWins, ",")(Index)) Then
                    'intCurPage = Val(Mid(Me.vfgCallingData.TextMatrix(vfgCallingData.Rows - 1, i), 4, InStr(1, Me.vfgCallingData.TextMatrix(vfgCallingData.Rows - 1, i), "/")))
                    intCurPage = 1
                    Exit For
                End If
            Next

            '����¼���Ƶ���ǰҳ����λ��
'            For i = 1 To mIntPraRow * mRowRec * (intCurPage - 1)
'                If Not mrsPreparingData(Index).EOF Then
'                    mrsPreparingData(Index).MoveNext
'                End If
'            Next

        Else
            intCurPage = 1
        End If
        
        count = 0
        i = 3
        intcol = mRowRec * Index
        'ѭ����¼������������ʾ������
        Do While Not mrsPreparingData(Index).EOF
            If Nvl(mrsPreparingData(Index)!����) <> "" Then
                count = count + 1
                '�������������ʾԤ���������˳�ѭ��
                If count > mType_para.int����ҩ���� Then
                    Exit Do
                End If
                Me.vfgCallingData.Cell(flexcpFontSize, i, intcol, i, intcol) = Val(GetSetting("ZLSOFT", strReg, "�ֺ�(3)", "14"))
                Me.vfgCallingData.Cell(flexcpForeColor, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "����ҩ��ɫ", vbGreen)
                Me.vfgCallingData.Cell(flexcpFontName, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "����(3)", "����")
                Me.vfgCallingData.Cell(flexcpFontBold, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "����(3)", "false")
                Me.vfgCallingData.Cell(flexcpFontItalic, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "б��(3)", "false")
                Me.vfgCallingData.TextMatrix(i, intcol) = Nvl(mrsPreparingData(Index)!����)
                'һ�����ݼ�����֮��������һ��
                If intcol < mType_para.int����ҩ���� + Index * mRowRec - 1 Then
                    intcol = intcol + 1
                Else
                    intcol = mRowRec * Index
                    i = i + 1
                End If
'                If count = mRowRec Then
'                    count = 0
'                    intcol = Index * mRowRec
'                End If
                'mstrPraNames(Index) = mstrPraNames(Index) & Nvl(mrsPreparingData(Index)!����) & ","
            End If
            mrsPreparingData(Index).MoveNext
            
            'һ�����ݼ�����֮��������һ��
'            intcol = intcol + 1
'            If count = mRowRec Then
'                count = 0
'                intcol = Index * mRowRec
'                i = i + 1
'            End If
        Loop
        
        '��ʾ��ҳ��Ϣ
'        intcol = Index * mRowRec
'        For intcol = Index * mRowRec To (Index + 1) * mRowRec - 1
'            If (intcol \ mRowRec) Mod 2 = 0 Then
'                strTemp = String(0, " ")
'            Else
'                strTemp = String(1, " ")
'            End If
'
'            'Me.vfgCallingData.Cell(flexcpText, mintRows - 1, intcol, mintRows - 1, (Index + 1) * mRowRec - 1) = "����ҩ   " & strTemp & intCurPage & "/" & intPraPage & " ��" & mrsPreparingData(Index).RecordCount & "��"
'        Next
'        '�ϲ���ʾ��ҳ��Ϣ
'        vfgCallingData.MergeRow(Me.vfgCallingData.Rows - 1) = True
    End If
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************ShowPra*************")
        Call SaveErrLog("������ʱ��:" & gobjDatabase.Currentdate)
        Call SaveErrLog("������Ϣ:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub
Private Sub ShowTimeout(ByVal Index As Integer, ByVal intPage As Integer)
    Dim count As Integer
    Dim i As Integer
    Dim intPraPage As Integer
    Dim intCurPage As Integer
    Dim intcol As Integer
    Dim strTemp As String
    Dim strReg As String
    On Error GoTo errHandle
    strReg = "����ģ��\ҩ���Ŷӽк�\Һ������"
    If mType_para.bln��ʾ�ѹ��� Then
        '�����ѹ�����Ϣ
        loadTimeout (Split(mstrWins, ",")(Index)), Index
        intPraPage = 1
        If intPraPage = 0 Then intPraPage = 1
    
        '�ж��Ƿ��Ǵ������
        If intPage <> 1 Then
            intCurPage = 1
'            For i = 0 To Me.vfgCallingData.Cols - 1
'                '�õ���ǰҳ��
'                If vfgCallingData.TextMatrix(0, i) = (Split(mstrWins, ",")(Index)) Then
'                    intCurPage = Val(Mid(Me.vfgCallingData.TextMatrix(vfgCallingData.Rows - 1, i), 4, InStr(1, Me.vfgCallingData.TextMatrix(vfgCallingData.Rows - 1, i), "/")))
'                    intCurPage = 1
'                    Exit For
'                End If
'            Next
            '����¼���Ƶ���ǰҳ����λ��
'            For i = 1 To mType_para.int�ѹ������� * (intCurPage - 1)
'                If Not mrsPreparingData(Index).EOF Then
'                    mrsPreparingData(Index).MoveNext
'                End If
'            Next
        Else
            intCurPage = 1
        End If
        count = 0
        i = 3
        intcol = mRowRec * Index + IIf(mType_para.bln��ʾ����ҩ, mType_para.int����ҩ����, 0) + IIf(mType_para.bln��ʾ����ҩ, mType_para.int����ҩ����, 0)
        intPraPage = -1 * Int(-1 * mrsTimeoutData(Index).RecordCount / mType_para.int�ѹ�������)
        If mType_para.lng��ǰ����ҳ�� < intPraPage Then
            mType_para.lng��ǰ����ҳ�� = mType_para.lng��ǰ����ҳ�� + 1
        Else
            mType_para.lng��ǰ����ҳ�� = 1
        End If
        'ѭ����¼������������ʾ������
        mrsTimeoutData(Index).MoveFirst
        Do While Not mrsTimeoutData(Index).EOF
            If Nvl(mrsTimeoutData(Index)!����) <> "" Then
                count = count + 1
                '�������������ʾԤ���������˳�ѭ��
                If count > mType_para.int�ѹ������� * mType_para.lng��ǰ����ҳ�� Then
                    Exit Do
                End If
                If count <= mType_para.int�ѹ������� * mType_para.lng��ǰ����ҳ�� And count > mType_para.int�ѹ������� * (mType_para.lng��ǰ����ҳ�� - 1) Then
                    Me.vfgCallingData.Cell(flexcpFontSize, i, intcol, i, intcol) = Val(GetSetting("ZLSOFT", strReg, "�ֺ�(5)", "14"))
                    Me.vfgCallingData.Cell(flexcpForeColor, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "�ѹ�����ɫ", vbGreen)
                    Me.vfgCallingData.Cell(flexcpFontName, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "����(5)", "����")
                    Me.vfgCallingData.Cell(flexcpFontBold, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "����(5)", "false")
                    Me.vfgCallingData.Cell(flexcpFontItalic, i, intcol, i, intcol) = GetSetting("ZLSOFT", strReg, "б��(5)", "false")
                    Me.vfgCallingData.TextMatrix(i, intcol) = Nvl(mrsTimeoutData(Index)!����)
                    'һ�����ݼ�����֮��������һ��
                    If intcol < IIf(mType_para.bln��ʾ����ҩ, mType_para.int����ҩ����, 0) + IIf(mType_para.bln��ʾ����ҩ, mType_para.int����ҩ����, 0) + mType_para.int�ѹ������� + Index * mRowRec - 1 Then
                        intcol = intcol + 1
                    Else
                        intcol = mRowRec * Index + IIf(mType_para.bln��ʾ����ҩ, mType_para.int����ҩ����, 0) + IIf(mType_para.bln��ʾ����ҩ, mType_para.int����ҩ����, 0)
                        i = i + 1
                    End If
                End If
            End If
            mrsTimeoutData(Index).MoveNext
        Loop
    End If
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then   '3021��BOF �� EOF ��һ��Ϊ��
        Call SaveErrLog("***************ShowTimeout*************")
        Call SaveErrLog("������ʱ��:" & gobjDatabase.Currentdate)
        Call SaveErrLog("������Ϣ:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub

Public Sub SetFont()
    Dim strReg As String
    On Error GoTo errHandle
    strReg = "����ģ��\ҩ���Ŷӽк�\Һ������"
    With Me.vfgCallingData
        If .Cols = 0 Then Exit Sub
        '���������������ɫ��С
        '�к�����
        .Cell(flexcpFontSize, 1, 0, 1, .Cols - 1) = Val(GetSetting("ZLSOFT", strReg, "�ֺ�(0)", "14"))
        .Cell(flexcpForeColor, 1, 0, 1, .Cols - 1) = GetSetting("ZLSOFT", strReg, "��������ɫ", vbGreen)
        .Cell(flexcpFontName, 1, 0, 1, .Cols - 1) = GetSetting("ZLSOFT", strReg, "����(0)", "����")
        .Cell(flexcpFontBold, 1, 0, 1, .Cols - 1) = GetSetting("ZLSOFT", strReg, "����(0)", "false")
        .Cell(flexcpFontItalic, 1, 0, 1, .Cols - 1) = GetSetting("ZLSOFT", strReg, "б��(0)", "false")
        '��ҩ����ҩ�����ű�ͷ--�봰�ں����弰��ɫһ��
        .Cell(flexcpFontSize, 2, 0, 0, .Cols - 1) = Val(GetSetting("ZLSOFT", strReg, "�ֺ�(1)", "14"))
        .Cell(flexcpForeColor, 2, 0, 0, .Cols - 1) = GetSetting("ZLSOFT", strReg, "������ɫ", vbGreen)
        .Cell(flexcpFontName, 2, 0, 0, .Cols - 1) = GetSetting("ZLSOFT", strReg, "����(1)", "����")
        .Cell(flexcpFontBold, 2, 0, 0, .Cols - 1) = GetSetting("ZLSOFT", strReg, "����(1)", "false")
        .Cell(flexcpFontItalic, 2, 0, 0, .Cols - 1) = GetSetting("ZLSOFT", strReg, "б��(1)", "false")
        If mType_para.bln��ʾ���� Then
            .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = Val(GetSetting("ZLSOFT", strReg, "�ֺ�(1)", "14"))
            .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = GetSetting("ZLSOFT", strReg, "������ɫ", vbGreen)
            .Cell(flexcpFontName, 0, 0, 0, .Cols - 1) = GetSetting("ZLSOFT", strReg, "����(1)", "����")
            .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = GetSetting("ZLSOFT", strReg, "����(1)", "false")
            .Cell(flexcpFontItalic, 0, 0, 0, .Cols - 1) = GetSetting("ZLSOFT", strReg, "б��(1)", "false")
        End If
'        If mType_para.bln��ʾ����ҩ = True Then
'            .Cell(flexcpFontSize, 3, 0, mintRows - 1, 1) = Val(GetSetting("ZLSOFT", strReg, "�ֺ�(3)", "14"))
'            .Cell(flexcpForeColor, 3, 0, mintRows - 1, 1) = GetSetting("ZLSOFT", strReg, "����ҩ��ɫ", vbGreen)
'            .Cell(flexcpFontName, 3, 0, mintRows - 1, 1) = GetSetting("ZLSOFT", strReg, "����(3)", "����")
'            .Cell(flexcpFontBold, 3, 0, mintRows - 1, 1) = GetSetting("ZLSOFT", strReg, "����(3)", "false")
'            .Cell(flexcpFontItalic, 3, 0, mintRows - 1, 1) = GetSetting("ZLSOFT", strReg, "б��(3)", "false")
'        End If
'        If mType_para.bln��ʾ����ҩ = True Then
'            .Cell(flexcpFontSize, 3, 2, mintRows - 1, 3) = Val(GetSetting("ZLSOFT", strReg, "�ֺ�(2)", "14"))
'            .Cell(flexcpForeColor, 3, 2, mintRows - 1, 3) = GetSetting("ZLSOFT", strReg, "����ҩ��ɫ", vbGreen)
'            .Cell(flexcpFontName, 3, 2, mintRows - 1, 3) = GetSetting("ZLSOFT", strReg, "����(2)", "����")
'            .Cell(flexcpFontBold, 3, 2, mintRows - 1, 3) = GetSetting("ZLSOFT", strReg, "����(2)", "false")
'            .Cell(flexcpFontItalic, 3, 2, mintRows - 1, 3) = GetSetting("ZLSOFT", strReg, "б��(2)", "false")
'        End If
'        If mType_para.bln��ʾ�ѹ��� = True Then
'            .Cell(flexcpFontSize, 3, 4, mintRows - 1, 4) = Val(GetSetting("ZLSOFT", strReg, "�ֺ�(5)", "14"))
'            .Cell(flexcpForeColor, 3, 4, mintRows - 1, 4) = GetSetting("ZLSOFT", strReg, "�ѹ�����ɫ", vbGreen)
'            .Cell(flexcpFontName, 3, 4, mintRows - 1, 4) = GetSetting("ZLSOFT", strReg, "����(5)", "����")
'            .Cell(flexcpFontBold, 3, 4, mintRows - 1, 4) = GetSetting("ZLSOFT", strReg, "����(5)", "false")
'            .Cell(flexcpFontItalic, 3, 4, mintRows - 1, 4) = GetSetting("ZLSOFT", strReg, "б��(5)", "false")
'        End If
    End With
    
    Me.lblmsg.ForeColor = GetSetting("ZLSOFT", strReg, "����������ɫ", vbBlack)
    Me.lblmsg.FontSize = Val(GetSetting("ZLSOFT", strReg, "�ֺ�(4)", "14"))
    Me.lblmsg.FontName = GetSetting("ZLSOFT", strReg, "����(4)", "����")
    Me.lblmsg.FontBold = GetSetting("ZLSOFT", strReg, "����(4)", "false")
    Me.lblmsg.FontItalic = GetSetting("ZLSOFT", strReg, "б��(4)", "false")
    Exit Sub
errHandle:
    If err.Number > 0 And err.Number <> 3021 Then
        Call SaveErrLog("***************SetFont*************")
        Call SaveErrLog("������ʱ��:" & gobjDatabase.Currentdate)
        Call SaveErrLog("������Ϣ:" & err.Number & "->" & err.Description)
        Call SaveErrLog("************************************")
    End If
    Resume Next
End Sub
Private Function TranRowNum(ByVal lngRowNum As Long) As String
    '��ȦȦ��������ʾ��ʽ,���ֻ֧��1-10
    '  1-9 TranRowNum = Chr(Asc(lngRowNum) - 23896)
    '  10  TranRowNum = Chr(Asc(lngRowNum) - 23887)
    TranRowNum = ""
    If mType_para.bln��ʾ��ҩ��� Then
        If mType_para.int����ҩ���� <= 10 Then
            If lngRowNum < 10 Then
                TranRowNum = Chr(Asc(lngRowNum) - 23896)
            Else
                TranRowNum = Chr(Asc(lngRowNum) - 23887)
            End If
        Else
            TranRowNum = lngRowNum & "."
        End If
    End If
End Function

'Private Sub GetTotal(ByVal intType As Integer)
'    Dim i As Integer
'    Dim strTemp As String
'
'    If intType = 0 Then
'        For i = 0 To mintCols - 1
'            strTemp = ","
'            If Not mrsData(i) Is Nothing Then mrsData(i).MoveFirst
'            Do While mrsData(i).EOF
'
'                If InStr(1, strTemp, "," & mrsData(i)!���� & ",") Then
'                    mintSenpages(i) = mintSenpages(i) + 1
'                End If
'                strTemp = strTemp & mrsData(i)!���� & ","
'                mrsData(i).MoveNext
'            Loop
'            mrsData(i).MoveFirst
'        Next
'    Else
'
'    End If
'End Sub


