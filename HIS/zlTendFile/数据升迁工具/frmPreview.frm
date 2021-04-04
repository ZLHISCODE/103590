VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPreview 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "preView"
   ClientHeight    =   5130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form24"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtLength 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   1005
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   90
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1995
      LargeChange     =   10
      Left            =   3540
      Max             =   100
      SmallChange     =   2
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   285
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2685
      Left            =   90
      ScaleHeight     =   2655
      ScaleWidth      =   3225
      TabIndex        =   0
      Top             =   630
      Width           =   3255
      Begin VSFlex8Ctl.VSFlexGrid VsfData 
         Height          =   1455
         Left            =   570
         TabIndex        =   5
         Top             =   930
         Width           =   2265
         _cx             =   3995
         _cy             =   2566
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
         BackColorSel    =   16764057
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   3
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   5000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPreview.frx":0000
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
         AutoSizeMouse   =   0   'False
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
      Begin VB.Label lblDownTable 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������ɻ���"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   270
         TabIndex        =   4
         Top             =   1020
         Width           =   1125
      End
      Begin VB.Label lblUpTable 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������ɻ���"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   270
         TabIndex        =   3
         Top             =   600
         Width           =   1125
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1380
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.Line lineRight 
         X1              =   1380
         X2              =   1380
         Y1              =   360
         Y2              =   2220
      End
      Begin VB.Line lineLeft 
         X1              =   720
         X2              =   720
         Y1              =   360
         Y2              =   2220
      End
      Begin VB.Line lineBottom 
         X1              =   630
         X2              =   2790
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line lineTop 
         X1              =   630
         X2              =   2790
         Y1              =   600
         Y2              =   600
      End
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objStream As TextStream
Dim lngFormat As Long               '��ʽID
Dim lngFile As Long                 '���˻����ļ�.ID
Dim mlngRows As Long
Dim mstrSQL As String
Dim mstrSQL�� As String
Dim mstrSQL�� As String
Dim mstrSQL�� As String
Dim mstrSQL���� As String

'�����ļ���ʽ�������
Private mintTabTiers As Integer     '��ͷ���
Private mintTagFormHour As Integer  '��ʼʱ������
Private mintTagToHour As Integer    '��ֹʱ������
Private mobjTagFont As New StdFont  '������ʽ����
Private mlngTagColor As Long        '������ʽ��ɫ
Private mstrPaperSet As String      '��ʽ
Private mstrPageHead As String      'ҳü
Private mstrPageFoot As String      'ҳ��
Private mblnChildForm As Boolean
Private mlngActiveRows As Long      '��Ч������
Private mstrSubhead As String       '���ϱ�ǩ
Private mstrTabHead As String       '��ͷ��Ԫ
Private mstrPreHead As String       '�账�����,�ı�����Ŀ�����л�󶨶����Ŀ����
Private mstrColWidth As String      '�п����д�
Private mstrColumns As String       '��ǰ�����ļ����ж�Ӧ����Ŀ
Private lngCurColor As Long, strCurFont As String, objFont As StdFont
Private mrsItems As New ADODB.Recordset

Dim dblTitle As Double      '�������ĸ߶�
Dim dblUpTable As Double    '������ĸ߶�
Dim dblDownTable As Double  '������ĸ߶�

Private Const EM_GETLINECOUNT = &HBA&        '��ȡ������
Private Const EM_GETLINE = &HC4&             '����һ���ı���ָ����������
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub Form_Load()
    Dim lngRows As Long                             '����������Ч����
    Dim lngFixRows As Long                          '�̶�����
    Dim dblRowHeight As Double                      '�и�
    Dim lngParent As Long
    Dim strUpText As String
    Dim lngHeight As Long, lngWidth As Long         '��Ч�߶ȣ����
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    Dim rsTemp As New ADODB.Recordset
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = Screen.Height
    Me.Width = Screen.Width
    
    '����ҳ���ʽ
    Call zlGetPrinterSet
    
    '��ȡ��ӡ����ǰ״̬
    picDraw.Height = Printer.Height
    picDraw.Width = Printer.Width
    picDraw.ScaleHeight = Printer.ScaleHeight
    picDraw.ScaleWidth = Printer.ScaleWidth
    'ҳ�߾�
    lngTop = marrFormat(6)
    lngBottom = marrFormat(7)
    lngLeft = marrFormat(4)
    lngRight = marrFormat(5)
    'ʵ����Ч�߶ȣ����
    lngHeight = picDraw.ScaleHeight - lngTop - lngBottom
    lngWidth = picDraw.ScaleWidth - lngLeft - lngRight
    
    '��,�±߾�(lngTop , lngBottom)
    '��,�ұ߾�(lngLeft , lngRight)
    lineTop.X1 = 0
    lineTop.X2 = picDraw.ScaleWidth
    lineTop.Y1 = lngTop
    lineTop.Y2 = lngTop
    lineBottom.X1 = 0
    lineBottom.X2 = picDraw.ScaleWidth
    lineBottom.Y1 = picDraw.ScaleHeight - lngBottom
    lineBottom.Y2 = lineBottom.Y1
    
    lineLeft.X1 = lngLeft
    lineLeft.X2 = lngLeft
    lineLeft.Y1 = 0
    lineLeft.Y2 = picDraw.ScaleHeight
    lineRight.X1 = picDraw.ScaleWidth - lngRight
    lineRight.X2 = lineRight.X1
    lineRight.Y1 = 0
    lineRight.Y2 = picDraw.ScaleHeight
    
    '׼������������,����������,���±�ǩ������
    gstrSQL = "" & _
            "SELECT id, �ļ�id, nvl(��id,0) ��ID, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���, " & vbNewLine & _
            "       Ԥ�����id, �������, ʹ��ʱ��, ����Ҫ��id, �滻��, Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��" & vbNewLine & _
            "FROM �����ļ��ṹ A" & vbNewLine & _
            "WHERE A.�ļ�ID=[1]" & vbNewLine & _
            "ORDER BY A.��ID,A.�������"
    Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡ�����ļ�����", lngFormat)
    
    With rsTemp
        '1�����������ϱ߾࿪ʼ
        .Filter = "�������=1 And ��ID=0"
        lngParent = !ID
        '�и�
        .Filter = "��ID=" & lngParent & " And �������=3"
        dblRowHeight = !�����ı�
        '�̶�����
        .Filter = "��ID=" & lngParent & " And �������=1"
        lngFixRows = !�����ı�
        '����������������
        .Filter = "��ID=" & lngParent & " And �������=8"
        lblTitle.FontName = Split(!�����ı�, ",")(0)
        lblTitle.FontSize = Split(!�����ı�, ",")(1)
        .Filter = "��ID=" & lngParent & " And �������=7"
        lblTitle.Caption = !�����ı�
        '���������������
        .Filter = "��ID=" & lngParent & " And �������=4"
        lblUpTable.FontName = Split(!�����ı�, ",")(0)
        lblUpTable.FontSize = Split(!�����ı�, ",")(1)
        .Filter = "��ID=" & lngParent & " And �������=5"
        lblUpTable.BackColor = !�����ı�
        
        '���ñ���������
        picDraw.FontName = lblTitle.FontName
        picDraw.FontSize = lblTitle.FontSize
        lblTitle.Left = lngLeft
        lblTitle.Top = lngTop + 30
        lblTitle.Width = lngWidth
        lblTitle.Height = picDraw.TextHeight("a")
        
        '2�����ϱ�ǩ�ӱ������¿�ʼ
        .Filter = "�������=2 And ��ID=0"
        lngParent = !ID
        .Filter = "��ID=" & lngParent
        Do While Not .EOF
            strUpText = strUpText & IIf(strUpText = "", "", "  ") & IIf(!�Ƿ��� = 0, "", vbCrLf) & NVL(!�����ı�) & !Ҫ������
            .MoveNext
        Loop
        If strUpText <> "" Then
            lblUpTable.Caption = strUpText
            lblUpTable.AutoSize = True
        End If
        '���ñ���������
        picDraw.FontName = lblUpTable.FontName
        picDraw.FontSize = lblUpTable.FontSize
        lblUpTable.Left = lngLeft
        lblUpTable.Top = lblTitle.Top + lblTitle.Height + 30
        lblUpTable.Width = picDraw.ScaleWidth
        
        '3�����ñ��
    lngHeight = lngHeight - lblUpTable.Height - lblTitle.Height
        VsfData.Top = lblUpTable.Top + lblUpTable.Height + 30
        VsfData.Left = lngLeft
        VsfData.Width = lngWidth
        lngHeight = lngHeight + lngTop - VsfData.Top
        VsfData.Height = lngHeight
    lngRows = CLng(lngHeight \ dblRowHeight) - lngFixRows
        VsfData.Rows = lngFixRows + lngRows
        VsfData.FixedRows = lngFixRows
        VsfData.RowHeightMin = dblRowHeight
        
        mlngRows = lngRows
    End With
    
    Call VScroll1_Change
    
    If mrsItems.State = 0 Then
        '���ִ��ڵ����л����¼��Ŀ
        gstrSQL = " Select ��Ŀ���,��Ŀ����,��Ŀ����,��Ŀ����,��Ŀ����,��ĿС��,��Ŀ��ʾ,��Ŀ��λ,��Ŀֵ��,����ȼ�,Ӧ�÷�ʽ" & _
                  " From �����¼��Ŀ B" & _
                  " Where B.Ӧ�÷�ʽ<>0 " & _
                  " Order by ��Ŀ���"
        Set mrsItems = OpenSQLRecord(gstrSQL, "���ִ��ڵ����л����¼��Ŀ")
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload Me
End Sub

Private Sub VScroll1_Change()
    picDraw.Top = -1 * VScroll1.Value * (picDraw.Height - Me.Height) / 100
End Sub

Public Function ShowMe(ByVal objParent As Object, ByVal lngFileID As Long, arrData, Optional ByVal blnHide As Boolean = True) As Long
    '��ȡ�����¼���ĸ�ʽ
    mlngRows = 0
    lngFormat = lngFileID
    marrFormat = arrData
    If blnHide Then
        Unload frmPreview
        Load frmPreview
    Else
        Me.Show 1, objParent
    End If
    ShowMe = mlngRows
End Function

Public Function AnaliseData(ByVal objParent As Object, ByVal lngFileID As Long, arrData, objStream_ As TextStream) As Boolean
    lngFormat = lngFileID
    marrFormat = arrData
    Set objStream = objStream_
    
    Unload frmPreview
    Load frmPreview
    
    If Not ReadStruDef Then
        'û����Ҫ��������,���ֱ�ӷ��ؽ����ɹ�,Ӧ�ò������������
        AnaliseData = (mstrPreHead = "")
        Exit Function
    End If
    If Not ReadData Then Exit Function
    
    AnaliseData = True
End Function

Private Function ReadData() As Boolean
    Dim rsPati As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '��ȡ����ʹ�øü�¼�ļ��Ĳ����б�
    gstrSQL = "Select ID,����ID,����ID,��ҳID,Ӥ�� from ���˻����ļ� where nvl(����,0)=0 And ��ʽID=[1]"
    Set rsPati = OpenSQLRecord(gstrSQL, "��ȡʹ�øû����ļ��Ĳ����б�", lngFormat)
    Do While Not rsPati.EOF
        'װ������
        lngFile = rsPati!ID
        gstrSQL = mstrSQL
        Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡ��������", CLng(rsPati!ID), CLng(rsPati!����ID), CLng(rsPati!��ҳID), CLng(rsPati!Ӥ��))
        '�����ݲ����û����¼���ĸ�ʽ,ͬʱʵ��һ�����ݷ�����ʾ�Ĺ���
        Call PreTendFormat(rsTemp)
        '����ÿ������
        If Not ParseData Then Exit Function
        
        gcnOracle.Execute "ZL_���˻����ļ�_����(" & rsPati!ID & ")", , adCmdStoredProc
        'objStream.WriteLine "�ļ�ID:" & rsPati!ID & "������ID=" & rsPati!����ID & ";��ҳID=" & rsPati!��ҳID & ";Ӥ��=" & rsPati!Ӥ�� & "��" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "���..."
        
        If gintAutoRUN = 1 Then
            If Format(Now, "HH:mm") >= gstrEndTime Then
                Exit Do
            End If
        End If
        rsPati.MoveNext
    Loop
    
    ReadData = True
    Exit Function
errHand:
    MsgBox Err.Description
End Function

Private Function ParseData() As Boolean
    Dim arrCol, arrData
    Dim blnNewPage As Boolean
    Dim lngMutilRow As Long, lngRecord As Long
    Dim lngRow As Long, lngCount As Long
    Dim lngCol As Long, lngMAX As Long
    Dim lngStartPage As Long, lngEndPage As Long, lngStartRow As Long, lngEndRow As Long
    On Error GoTo errHand
    'ѭ����������������(һ�а󶨶����Ŀ,������ĿΪ�ı���)
    
    arrCol = Split(mstrPreHead, ",")
    lngMAX = UBound(arrCol)
    lngCount = VsfData.Rows - 1
    lngStartPage = 1: lngEndPage = 1: lngStartRow = 1: lngEndRow = 1
    
    For lngRow = 1 To lngCount
        lngMutilRow = 0
        lngRecord = Val(VsfData.TextMatrix(lngRow, VsfData.Cols - 1))
        If lngRecord <> 0 Then
            For lngCol = 0 To lngMAX
                If VsfData.TextMatrix(lngRow, arrCol(lngCol)) <> "" Then
                '׼����ֵ
                With txtLength
                    .Width = VsfData.ColWidth(arrCol(lngCol))
                    .Text = VsfData.TextMatrix(lngRow, arrCol(lngCol))
                    .FontName = VsfData.FontName
                    .FontSize = VsfData.FontSize
                End With
                arrData = GetData(txtLength.Text)
                If UBound(arrData) > lngMutilRow Then lngMutilRow = UBound(arrData)
                End If
            Next
            
            lngEndRow = (lngStartRow + lngMutilRow)
reSub:
            If lngEndRow > mlngActiveRows Then
                blnNewPage = True
                lngEndRow = lngEndRow - mlngActiveRows
                lngEndPage = lngEndPage + 1
                GoTo reSub
            End If
            
            'һ�н���ʱ��������ӡ��������
            gstrSQL = "ZL_���˻����ӡ��Ǩ_UPDATE(" & lngRecord & "," & lngFile & "," & lngMutilRow + 1 & "," & lngStartPage & "," & lngStartRow & "," & lngEndPage & "," & lngEndRow & ")"
            gcnOracle.Execute gstrSQL, , adCmdStoredProc
            
            If blnNewPage Then
                lngStartPage = lngEndPage
                blnNewPage = False
            End If
            lngStartRow = lngEndRow + 1
            If lngStartRow > mlngActiveRows Then
                lngStartRow = lngStartRow - mlngActiveRows
                lngStartPage = lngStartPage + 1
                If lngEndPage < lngStartPage Then lngEndPage = lngStartPage
            End If
        End If
    Next
    
    ParseData = True
    Exit Function
errHand:
    MsgBox Err.Description
End Function

Private Sub PreTendFormat(ByVal rsTemp As ADODB.Recordset)
    Dim blnTag As Boolean
    Dim aryItem() As String
    Dim lngRow As Long, lngCol As Long, lngCount As Long, strCell As String
    On Error GoTo errHand
    
    '���û����¼���ĸ�ʽ
    With VsfData
        .FixedRows = 3
        .Clear
        Set .DataSource = rsTemp
        
        '��ͷ��д
        .MergeCells = flexMergeFixedOnly
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True
        .ColHidden(.Cols - 1) = True
        .ColHidden(.Cols - 2) = True
        .ColHidden(.Cols - 3) = True
        
        '������ͷ
        aryItem = Split(mstrTabHead, "|")
        For lngCount = 0 To UBound(aryItem)
            strCell = aryItem(lngCount)
            lngRow = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            lngCol = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            .TextMatrix(lngRow, lngCol + 1) = strCell
        Next
        
        '�п�����
        Dim blnAlign As Boolean
        aryItem = Split(mstrColWidth, ",")
        For lngCount = 2 To .Cols - 1
            If Not .ColHidden(lngCount) Then
                .ColWidth(lngCount) = Val(Split(aryItem(lngCount - 2), "`")(0))
                If InStr(1, aryItem(lngCount - 2), "`") <> 0 Then
                    blnAlign = True
                    .ColAlignment(lngCount) = Val(Split(aryItem(lngCount - 2), "`")(1))
                End If
            End If
        Next
        
        '�̶��и�ʽΪ����
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        '�ٰ��кϲ�
        For lngCount = 0 To VsfData.Cols - 1
            VsfData.MergeCol(lngCount) = True
        Next
        .AutoSize 0, .Cols - 1
        
        If blnAlign = False Then
            '��Ϊ�����û���������ʾ�ж��뷽ʽ
            If .FixedRows < .Rows Then .Cell(flexcpAlignment, .FixedRows, 0, .Rows - 1, .Cols - 1) = flexAlignGeneralCenter
        End If
        For lngCount = 0 To .Rows - 1
            If .RowHeight(lngCount) < .RowHeightMin Then .RowHeight(lngCount) = .RowHeightMin
        Next
        Select Case mintTabTiers
        Case 1
            .RowHidden(0) = False
            .RowHidden(1) = True
            .RowHidden(2) = True
        Case 2
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = True
        Case 3
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = False
        End Select
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next
        
    End With
    Exit Sub
errHand:
    MsgBox Err.Description
End Sub

Private Function ReadStruDef() As Boolean
    Dim arrCol
    Dim intCol As Integer, intCount As Integer
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '��ȡ�����ļ���ʽ����
    gstrSQL = "Select d.�������, d.�����ı�, d.Ҫ������" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '�����ʽ'" & _
        " Order By d.�������"
    Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡ�����ļ���ʽ����", lngFormat)
    With rsTemp
        Do While Not .EOF
            Select Case "" & !Ҫ������
            Case "��ͷ����": mintTabTiers = Val("" & !�����ı�)
            Case "������":  VsfData.Cols = Val("" & !�����ı�)
            Case "��С�и�": VsfData.RowHeightMin = Val("" & !�����ı�)
            Case "�ı�����"
                strCurFont = "" & !�����ı�
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set VsfData.Font = objFont
                Set lblUpTable.Font = VsfData.Font
                Set Font = lblUpTable.Font
                
            Case "�ı���ɫ": VsfData.ForeColor = Val("" & !�����ı�)
            Case "�����ɫ": VsfData.GridColor = Val("" & !�����ı�): VsfData.GridColorFixed = VsfData.GridColor
            
            Case "�����ı�": lblTitle.Caption = "" & !�����ı�
            Case "��������"
                strCurFont = "" & !�����ı�
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set lblTitle.Font = objFont
                lblTitle.AutoSize = False
            
            Case "��ʼʱ��": mintTagFormHour = Val("" & !�����ı�)
            Case "��ֹʱ��": mintTagToHour = Val("" & !�����ı�)
            Case "��������"
                strCurFont = "" & !�����ı�
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set mobjTagFont = objFont
            Case "������ɫ": mlngTagColor = Val("" & !�����ı�)
            Case "��Ч������": mlngActiveRows = Val(!�����ı�)
            End Select
            .MoveNext
        Loop
    End With
    
    gstrSQL = "Select ��ʽ, ҳü, ҳ��,���� From ����ҳ���ʽ Where ���� = 3 And ��� In (Select ҳ�� From �����ļ��б� Where Id = [1])"
    Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡ����ҳ���ʽ", lngFormat)
    If Not rsTemp.EOF Then
        mstrPaperSet = "" & rsTemp!��ʽ: mstrPageHead = "" & rsTemp!ҳü: mstrPageFoot = "" & rsTemp!ҳ��
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.�������, d.�����ı�, d.Ҫ������, Nvl(d.�Ƿ���, 0) As �Ƿ���" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���ϱ�ǩ'" & _
        " Order By d.�������"
    Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡ���ϱ�ǩ����", lngFormat)
    With rsTemp
        mstrSubhead = ""
        Do While Not .EOF
            mstrSubhead = mstrSubhead & "|" & IIf(!�Ƿ��� = 0, "", vbCrLf) & !�����ı� & "{" & !Ҫ������ & "}"
            .MoveNext
        Loop
        If mstrSubhead <> "" Then mstrSubhead = Mid(mstrSubhead, 2)
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.�������, d.�����д�, d.�����ı�" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '��ͷ��Ԫ'" & _
        " Order By d.�������"
    Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡ��ͷ��Ԫ����", lngFormat)
    With rsTemp
        mstrTabHead = ""
        Do While Not .EOF
            mstrTabHead = mstrTabHead & "|" & !�����д� - 1 & "," & !������� & "," & !�����ı�
            .MoveNext
        Loop
        If mstrTabHead <> "" Then mstrTabHead = Mid(mstrTabHead, 2)
    End With
    
    '��ѯ�����֯
    '------------------------------------------------------------------------------------------------------------------
    Dim strSql�� As String, str��ʽ As String
    Dim bln���� As Boolean, blnʱ�� As Boolean, bln��ʿ As Boolean
    Dim blnǩ���� As Boolean, blnǩ��ʱ�� As Boolean, blnǩ������ As Boolean
    Dim lngColumn As Long
    
    gstrSQL = "Select d.�������, d.��������, d.�����д�, d.�����ı�, d.Ҫ������, d.Ҫ�ص�λ,d.Ҫ�ر�ʾ " & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���м���'" & _
        " Order By d.�������, d.�����д�"
    Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡ���м��϶���", lngFormat)
    With rsTemp
        lngColumn = 0: mstrColumns = "": mstrColWidth = ""
        mstrSQL�� = "": mstrSQL�� = "": strSql�� = "": mstrSQL�� = "": mstrSQL���� = ""
        bln���� = False: blnʱ�� = False: bln��ʿ = False
        blnǩ���� = False: blnǩ��ʱ�� = False: blnǩ������ = False
        Do While Not .EOF
            
            If lngColumn <> !������� Then
                mstrColumns = mstrColumns & IIf(mstrColumns = "", "", ";1;" & str��ʽ) & "|" & !������� & ";" & !Ҫ������
                mstrColWidth = mstrColWidth & "," & !��������
                str��ʽ = ""
                If !Ҫ������ <> "" Then
                    str��ʽ = "{" & NVL(!�����ı�) & "[" & !Ҫ������ & "]" & NVL(!Ҫ�ص�λ) & "}"
                    mstrSQL�� = mstrSQL�� & "," & Mid(strSql��, 3) & " As C" & Format(lngColumn, "00")
                Else
                    If strSql�� <> "" Then
                        mstrSQL�� = mstrSQL�� & "," & Mid(strSql��, 3) & " As C" & Format(lngColumn, "00")
                    Else
                        mstrSQL�� = mstrSQL�� & ",'' As C" & Format(lngColumn, "00")
                    End If
                End If
                strSql�� = ""
                lngColumn = !�������
            Else
                mstrColumns = mstrColumns & "," & !Ҫ������
                str��ʽ = str��ʽ & "{" & NVL(!�����ı�) & "[" & !Ҫ������ & "]" & NVL(!Ҫ�ص�λ) & "}"
            End If
            
            Select Case !Ҫ������
            Case "����"
                bln���� = True
                mstrSQL�� = mstrSQL�� & ",����"
                mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, 'yyyy-mm-dd') As ����"
                strSql�� = strSql�� & "||" & !Ҫ������
            Case "ʱ��"
                blnʱ�� = True
                mstrSQL�� = mstrSQL�� & ",ʱ��"
                mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, 'hh24:mi') As ʱ��"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "ǩ����"
                blnǩ���� = True
                mstrSQL�� = mstrSQL�� & ",ǩ����"
                mstrSQL�� = mstrSQL�� & ",l.ǩ���� As ǩ����"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "ǩ��ʱ��"
                blnǩ��ʱ�� = True
                mstrSQL�� = mstrSQL�� & ",ǩ��ʱ��"
                mstrSQL�� = mstrSQL�� & ",Decode(a.��Ŀ����,Null,Null,Substr(a.��Ŀ����,12,5)) As ǩ��ʱ��"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "ǩ������"
                blnǩ������ = True
                mstrSQL�� = mstrSQL�� & ",ǩ������"
                mstrSQL�� = mstrSQL�� & ",Decode(a.��Ŀ����,Null,Null,Substr(a.��Ŀ����, 1,11)) As ǩ������"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "��ʿ"
                bln��ʿ = True
                mstrSQL�� = mstrSQL�� & ",��ʿ"
                mstrSQL�� = mstrSQL�� & ",l.������ As ��ʿ"
                strSql�� = strSql�� & "||" & !Ҫ������
            Case Else
                If !Ҫ������ <> "" Then
                    mstrSQL�� = mstrSQL�� & ",Max(""" & !Ҫ������ & """) As """ & !Ҫ������ & """"
                    mstrSQL���� = mstrSQL���� & " Or """ & !Ҫ������ & """ Is Not Null"
                    strSql�� = strSql�� & "||""" & !Ҫ������ & """"
                    
                    If Trim("" & !�����ı�) = "" And Trim("" & !Ҫ�ص�λ) = "" Then
                        mstrSQL�� = mstrSQL�� & ", Decode(c.��Ŀ����, '" & !Ҫ������ & "', Nvl(c.δ��˵��,c.��¼����), '') As """ & !Ҫ������ & """"
                    Else
                        mstrSQL�� = mstrSQL�� & ", Decode(c.��Ŀ����, '" & !Ҫ������ & "', Nvl(c.δ��˵��,Decode(c.��¼����,Null,Null,'" & !�����ı� & "'||c.��¼����||'" & !Ҫ�ص�λ & "')), '') As """ & !Ҫ������ & """"
                    End If
                End If
            End Select
            .MoveNext
        Loop
        
        mstrColWidth = Mid(mstrColWidth, 2)
        '�������һ�еĸ�ʽ
        mstrColumns = mstrColumns & IIf(mstrColumns = "", "", ";1;" & str��ʽ) '& "|" & !������� & ";" & !Ҫ������
        mstrColumns = Mid(mstrColumns, 2)     '��ʽ��:�к�;��Ŀ����1,��Ŀ����2|�к�...,ʵ��;1;����|2;����|3...
        If Mid(strSql��, 3) <> "" Then
            mstrSQL�� = mstrSQL�� & "," & Mid(strSql��, 3) & " As C" & Format(lngColumn, "00")
        Else
            mstrSQL�� = mstrSQL�� & ",'' As C" & Format(lngColumn, "00")
        End If
        
        If mstrSQL���� <> "" Then mstrSQL���� = "(" & Mid(mstrSQL����, 5) & ")"
        
        '���û�г������ڣ�ʱ�䣬��ʿ�����ڲ���Ҫ���䣬�Ա�֤�в�����������
        If bln���� = False Then mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, 'yyyy-mm-dd') As ����"
        If blnʱ�� = False Then mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, 'hh24:mi') As ʱ��"
        If bln��ʿ = False Then mstrSQL�� = mstrSQL�� & ",l.������ As ��ʿ"
        
        If blnǩ���� = False Then mstrSQL�� = mstrSQL�� & ",l.ǩ���� As ǩ����"
        If blnǩ������ = False Then mstrSQL�� = mstrSQL�� & ",Decode(a.��Ŀ����,Null,Null,Substr(a.��Ŀ����,1,11)) As ǩ������"
        If blnǩ��ʱ�� = False Then mstrSQL�� = mstrSQL�� & ",Decode(a.��Ŀ����,Null,Null,Substr(a.��Ŀ����,12,5)) As ǩ��ʱ��"
        
        If Mid(mstrSQL��, 2) = "" Then
            MsgBox "�Բ�����û�ж��嵱ǰ��������ʾ����Ϣ�����ڲ����ļ������ж��壡"
            Exit Function
        End If
        
        '�����ڲ��������ӹ̶���
        mstrSQL�� = mstrSQL�� & ",MAX(֤��ID) AS ֤��ID,MAX(ǩ������) AS ǩ������,MAX(��¼ID) AS ��¼ID"
        mstrSQL�� = mstrSQL�� & ",A.��ĿID AS ֤��ID,NVL(A.��¼����,'��ʿ') AS ǩ������,C.��¼ID"
        mstrSQL�� = mstrSQL�� & ",֤��ID,ǩ������,��¼ID"
        
        '������Щ�е�������Ҫ���д�ӡ��������
        Dim arrData
        Dim strtodo As String
        Dim intto As Integer, intdo As Integer
        mstrPreHead = ""
        arrCol = Split(mstrColumns, "|")
        intCount = UBound(arrCol)
        For intCol = 0 To intCount
            If UBound(Split(Split(arrCol(intCol), ";")(3), "}{")) > 0 Then
                'ֻҪ��һ����������������Ϊ�ı��ʹ���
                
                strtodo = Split(arrCol(intCol), ";")(3)
                strtodo = Replace(strtodo, "]}{[", "||")
                strtodo = Replace(Replace(strtodo, "{[", ""), "]}", "")
                arrData = Split(strtodo, "||")
                intdo = UBound(arrData)
                For intto = 0 To intdo
                    mrsItems.Filter = "��Ŀ����='" & arrData(intto) & "'"
                    If mrsItems.RecordCount <> 0 Then
                        '����û�������Ŀʱ�������ó��ı���,��ô������20�����ϵ���Ŀ�ż��,�û����ý������͵����ó������Ͳ���ȷ
                        If mrsItems!��Ŀ���� = 1 And mrsItems!��Ŀ��ʾ = 0 And mrsItems!��Ŀ���� >= 20 Then
                            mstrPreHead = mstrPreHead & "," & Val(Split(arrCol(intCol), ";")(0)) + 1    '�����й̶����У�������Ŵ�0��ʼ�����+1
                            Exit For
                        End If
                    End If
                Next
            Else
                '����Ƿ�Ϊ�ı���
                mrsItems.Filter = "��Ŀ����='" & Replace(Replace(Split(arrCol(intCol), ";")(3), "{[", ""), "]}", "") & "'"
                If mrsItems.RecordCount <> 0 Then
                    '����û�������Ŀʱ�������ó��ı���,��ô������20�����ϵ���Ŀ�ż��,�û����ý������͵����ó������Ͳ���ȷ
                    If mrsItems!��Ŀ���� = 1 And mrsItems!��Ŀ��ʾ = 0 And mrsItems!��Ŀ���� >= 20 Then
                        mstrPreHead = mstrPreHead & "," & Val(Split(arrCol(intCol), ";")(0)) + 1    '�����й̶����У�������Ŵ�0��ʼ�����+1
                    End If
                End If
            End If
        Next
        
        mrsItems.Filter = 0
        If mstrPreHead = "" Then Exit Function
        mstrPreHead = Mid(mstrPreHead, 2)
        Call SQLCombination
    End With
    
    ReadStruDef = True
    Exit Function
errHand:
    MsgBox Err.Description
End Function

Private Sub SQLCombination()
    mstrSQL = "Select ����,����ʱ��," & Mid(mstrSQL��, 10) & vbCrLf & _
                " From (Select ��¼���,ʱ�� as ����,����ʱ��," & Mid(mstrSQL��, 2) & vbCrLf & _
                "        From (Select c.��¼���,l.����ʱ��," & Mid(mstrSQL��, 2) & vbCrLf & _
                "               From ���˻������� l, ���˻�����ϸ c,���˻�����ϸ a,���˻����ļ� f " & vbCrLf & _
                "               Where l.Id = c.��¼id And l.�ļ�ID=f.ID " & _
                "               And a.��¼id(+)=l.ID And a.��¼����(+)=5 And Nvl(a.��ֹ�汾,0)=0 And c.��ֹ�汾 Is Null And c.��¼����<>5  " & _
                "               And f.id=[1] And f.����id = [2] And f.��ҳid = [3] And Nvl(f.Ӥ��,0)=[4] )" & vbCrLf & _
                IIf(mstrSQL���� <> "", "Where " & mstrSQL����, "") & _
                "       Group By ����, ʱ��, ����ʱ��,��¼���,��ʿ,ǩ����,ǩ������,ǩ��ʱ��" & _
                                "       Order By ����, ʱ��, ����ʱ��,��¼���,��ʿ,ǩ����,ǩ������,ǩ��ʱ��)"
End Sub

'######################################################################################################################
'**********************************************************************************************************************
'��#�ָ��������ڵĴ��붼��������,û�±�
Private Function GetData(ByVal strInput As String) As Variant
    Dim arrData
    Dim strData As String
    Dim strLine(256) As Byte
    Dim lngRow As Long, lngRows As Long
    
    GetData = ""
    lngRows = SendMessage(txtLength.hWnd, EM_GETLINECOUNT, 0&, 0&)
    For lngRow = 1 To lngRows
        Call ClearArray(strLine)
        Call SendMessage(txtLength.hWnd, EM_GETLINE, lngRow - 1, strLine(0))
        strData = StrConv(strLine, vbUnicode)
        strData = TruncZero(strData)
        GetData = GetData & IIf(GetData = "", "", "|ZYB.ZLSOFT|") & strData
    Next
    GetData = Split(GetData, "|ZYB.ZLSOFT|")
End Function

Private Sub ClearArray(strLine() As Byte)
    Dim intdo As Integer, intMax As Integer
    intMax = UBound(strLine)
    For intdo = 0 To intMax
        strLine(intdo) = 0
    Next
    strLine(1) = 1
End Sub

Private Function TrimStr(ByVal str As String) As String
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�������ȥ�����˵Ŀո�

    If InStr(str, Chr(0)) > 0 Then
        TrimStr = Trim(Left(str, InStr(str, Chr(0)) - 1))
    Else
        TrimStr = Trim(str)
    End If
End Function

Private Function TruncZero(ByVal strInput As String) As String
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function
