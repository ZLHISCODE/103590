VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.UserControl ucScheduleTimetable 
   ClientHeight    =   6480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9240
   ScaleHeight     =   6480
   ScaleWidth      =   9240
   ToolboxBitmap   =   "ucScheduleTimetable.ctx":0000
   Begin VB.Timer timColor 
      Interval        =   100
      Left            =   480
      Top             =   6120
   End
   Begin VB.CommandButton btnSchLabel 
      Height          =   180
      Index           =   0
      Left            =   1080
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5520
      Width           =   90
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfTime 
      Height          =   5835
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8535
      _cx             =   15055
      _cy             =   10292
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
      Rows            =   25
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
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
   Begin VB.Menu menuTimePopup 
      Caption         =   "ʱ����Ҽ��˵�"
      Begin VB.Menu menuPopupTimeProjectAdd 
         Caption         =   "����ʱ��ƻ�"
      End
      Begin VB.Menu menuPopupTimeProjectModi 
         Caption         =   "�޸�ʱ��ƻ�"
      End
      Begin VB.Menu menuPopupTimeProjectDel 
         Caption         =   "ɾ��ʱ��ƻ�"
      End
      Begin VB.Menu menuPopupTimeSplit 
         Caption         =   "-"
      End
      Begin VB.Menu menuPopupTimeProjectColor 
         Caption         =   "ʱ�����ɫ����"
      End
   End
   Begin VB.Menu menuSchedulePopup 
      Caption         =   "ԤԼ��ǩ�Ҽ��˵�"
      Begin VB.Menu menuPopupScheduleModi 
         Caption         =   "�޸�ԤԼ"
      End
      Begin VB.Menu menuPopupSchedulePrint 
         Caption         =   "��ӡԤԼ��"
      End
   End
End
Attribute VB_Name = "ucScheduleTimetable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
'-------------------------------------------------------------
'   �ؼ��Ĺ����¼�
'-------------------------------------------------------------
Public Event OnMenuTimeProjectAdd()
Public Event OnMenuTimeProjectModify(ByVal lngTimeProjectID As Long)
Public Event OnMenuTimeProjectBeforeDel(ByRef blnCancel As Boolean)
Public Event OnMenuTimeProjectSetColor()
Public Event OnMenuScheduleModify()
Public Event OnMenuSchedulePrint()
Public Event OnChangeOrder(ByVal lngOrderID As Long, ByVal strOrderInfo As String)  '���ļ��ҽ���ı�ǩ
Public Event OnSchLabelModifed(ByVal iIndex As Integer)                             'ԤԼ��ǩ���ƶ������޸ĳ���
'-------------------------------------------------------------
'   �ؼ���˽������
'-------------------------------------------------------------
Private mblnDragingLabel As Boolean     '������קԤԼ��ǩ�ı��
Private mlngOriginLeft As Long          '��ǩ��ק֮ǰ����ǩ���е�һ����ǩ�����ڵ�Leftλ��
Private mlngOriginTop As Long           '��ǩ��ק֮ǰ����ǩ���е�һ����ǩ�����ڵ�Topλ��
Private mlngOriginMinute As Long        '��ǩ��ק֮ǰ��ʱ�䳤��
Private mlngBaseX As Long               '�ƶ���ǩ�Ļ�׼X�����ڱ�ǩ�飬����MouseDown��X
Private mlngBaseY As Long               '�ƶ���ǩ�Ļ�׼Y
Private mlngDownX As Long               'MouseDownʱ��X
Private mblnRestorePos As Boolean       'ԤԼ��ǩ�Ƿ���Ҫ�ص�ԭ����λ��

Private mblnSizingLabel As Boolean      '���ڵ���ԤԼ��ǩ�Ŀ��
Private mlngOriginWidth As Long         '����֮ǰ����ǰԤԼ��ǩ�Ŀ��

Private mlngColorTabRest As Long        'ԤԼʱ�����Ϣʱ����ɫ
Private mlngColorTabWork As Long        'ԤԼʱ�������ʱ����ɫ
Private mlngColorLblWaiting As Long     'ԤԼ��ǩ��ԤԼ�Ⱥ���ɫ
Private mlngColorLblDone As Long        'ԤԼ��ǩ�������ɫ
Private mlngColorLblPassed As Long      'ԤԼ��ǩ��������ɫ

Private mlngCurrTimePorjectID As Long   '��ǰ�������λ�õ�ʱ��ƻ�ID
Private mlngSchPlanID  As Long          '��ǰʱ����ԤԼ����ID
Private mlngSchDeviceID As Long         '��ǰԤԼ�豸ID
Private mlngOrderID As Long             '��ǰҽ��ID
Private mdtSchDate As Date              '��ǰ����
Private mblnMoved As Boolean            '�Ƿ��Ѿ�ת��
Private mstrOrderInfo As String         '��ǰҽ����ԤԼ��Ϣ
Private mColordict As New Dictionary
Private mlngCol As Long                 '��ǰ�������λ�õ��к�
Private mlngRow As Long                 '��ǰ�������λ�õ��к�

Private mlngMouseX As Long               '�������������ʱ�书��X����
Private mlngMouseY As Long               '�������������ʱ�书��Y����

Private mlngSchLabelIndex As Long       'ԤԼ��ǩ���������
Private mlngPoolIndex As Long           '��ǰѡ�е�ԤԼ������
Private mlngBtnIndex As Long            '��ǰѡ�е�ԤԼ��ǩ������ԤԼ����������һ��
Private mrsCalendar As ADODB.Recordset  '����ԤԼ�������ݼ�
Private mIsReadOnly As Boolean          '�Ƿ�ֻ��ģʽ
Private mstrModifiedOrderID As String   '�����ԤԼ��Ϣ��ҽ��ID�����á�,������
Private mlngFontSize As Long            '����ԤԼ��ǩ����Ĵ�С

Private mlngTimeProjectID As Long '��ǰʱ�������ʱ��ID,��������0

Private Const SchPreTime = 5            '����ԤԼ����ǰʱ�䣬5����
'ԤԼ�ؼ�ʹ��ģʽ
Private Enum constUseType
    Sch_UseType_���ԤԼ = 1            'ֻ�е�ǰ�½���ԤԼ��ǩ����ʹ��
    Sch_UseType_ԤԼ���� = 2            '����ԤԼ��ǩ������ʹ��
    Sch_UseType_ԤԼ���� = 3            '����ʾԤԼ��ǩ��ֻ��ʾʱ��ƻ�
End Enum
Private mlngUseType As constUseType     '��¼��ǰ��ԤԼģʽ

'���ԤԼ��������
Private Enum constSchedulePlanType
    Sch_PlanType_ÿ�� = 1
    Sch_PlanType_ÿ�� = 2
    Sch_PlanType_ÿ�� = 3
    Sch_PlanType_һ�� = 4
End Enum

'ԤԼ��ǩ����Ϣ
Private Type TYPE_SchLabel
    lng��� As Long             'ԤԼ���,ԤԼ�ذ������������
    lngRow As Long              '��ǩ���ڵ���
    lngCol As Long              '��ǩ���ڵ���
    lngҽ��ID As Long           'ҽ��ID
    strҽ������ As String       'ҽ������
    str���� As String           '����
    dtStartTime As Date         'ԤԼ��ʼʱ��
    dtEndTime As Date           'ԤԼ����ʱ��
    lngBtnCount As Long         'һ��ԤԼ��ǩ�����İ�ť�����������ǩ����2�����ϰ�ť
    lngBtnIndex As Long         'ԤԼ��ǩ��index������Ǳ�ǩ�飬��¼��һ����ǩ��index
    lngTimeProjectID As Long    'ʱ��ƻ�ID
    dt��ʼʱ��� As Date        '��ʼʱ���
    dt����ʱ��� As Date        '����ʱ���
    isModified As Boolean       '�Ƿ��Ѿ����޸�
    bln��ִ�� As Boolean        '�Ƿ��Ѿ�ִ�У�����ҽ������.ִ�й���=0 or =1 or =2, ִ�й���: -1-���أ�0��1-�ѵǼǣ�2-�ѱ�����3-�Ѽ�飻4-�ѱ��棻5-����ˣ�6-�����
    bln�ѱ��� As Boolean        '��ԤԼ�Ƿ��Ѿ����浽���ݿ�
End Type
Private mSchLabelPool() As TYPE_SchLabel

'ԤԼʱ��ƻ�
Private Type Type_TimeProject
    lngID As Long               'ʱ��ƻ�ID
    lngSchPlanID As Long        'ԤԼ����ID
    dtStartTime As Date         '��ʼʱ��
    dtEndTime As Date           '����ʱ��
    lngSum As Long              'ԤԼ����
    lngCalType As Long          '���㷽��
End Type
Private mSchTimeProject() As Type_TimeProject   '���տ�ʼʱ������
Private mlngSchSum As Long      'ԤԼ������

'-------------------------------------------------------------
'   �ؼ��Ĺ�������
'-------------------------------------------------------------
'����ֻ��ģʽ
Property Get IsReadOnly() As Boolean
    IsReadOnly = mIsReadOnly
End Property

Property Let IsReadOnly(value As Boolean)
    mIsReadOnly = value
    
    Call setReadOnly
End Property

'��ǰ��ѡ�е�ʱ��ƻ�ID
Property Get CurrTimeProjectID() As Long
     CurrTimeProjectID = mlngCurrTimePorjectID
End Property

'��ǰ��ԤԼ����ID
Property Get SchedulePlanID() As Long
    SchedulePlanID = mlngSchPlanID
End Property

'�ؼ���ʹ��ģʽ
Property Get UseType() As Long
    UseType = mlngUseType
End Property

'��ǰѡ��ԤԼ��ǩ�����
Property Get Label���() As Long
    Label��� = mSchLabelPool(mlngPoolIndex).lng���
End Property

'��ǰѡ��ԤԼ��ǩ�Ŀ�ʼʱ���
Property Get Label��ʼʱ���() As Date
    Label��ʼʱ��� = mSchLabelPool(mlngPoolIndex).dt��ʼʱ���
End Property

'��ǰѡ��ԤԼ��ǩ�Ľ���ʱ���
Property Get Label����ʱ���() As Date
    Label����ʱ��� = mSchLabelPool(mlngPoolIndex).dt����ʱ���
End Property

'��ǰѡ��ԤԼ��ǩ�Ŀ�ʼʱ��
Property Get Label��ʼʱ��() As Date
    Label��ʼʱ�� = mSchLabelPool(mlngPoolIndex).dtStartTime
End Property

'��ǰѡ��ԤԼ��ǩ�Ľ���ʱ��
Property Get Label����ʱ��() As Date
    Label����ʱ�� = mSchLabelPool(mlngPoolIndex).dtEndTime
End Property

'��ǰѡ��ԤԼ��ǩ��ҽ��ID
Property Get LabelOrderID() As Long
    LabelOrderID = mlngOrderID
End Property

'��ǰѡ��ԤԼ��ǩ��ȫ��ԤԼ��Ϣ
Property Get LabelOrderInfo() As String
    LabelOrderInfo = mstrOrderInfo
End Property

'��ǰѡ��ԤԼ��ǩ�Ļ�������
Property Get Label����() As String
    Label���� = mSchLabelPool(mlngPoolIndex).str����
End Property

'�����ԤԼ��Ϣ��ҽ��ID�����á�,������
Property Get strModifiedOrderID() As String
    If Trim(mstrModifiedOrderID) = "" Then
        strModifiedOrderID = mstrModifiedOrderID
    Else
        strModifiedOrderID = Mid(mstrModifiedOrderID, 2)
    End If
End Property

'-------------------------------------------------------------
'   �ؼ��Ĺ�������
'-------------------------------------------------------------
Public Function funSaveSchedule(dtSchStartTime As Date, dtSchEndTime As Date, lngOrderID As Long, _
    strName As String, lngNumber As Long, lngSchDeviceID As Long, dtSegStartTime As Date, _
    dtSegEndTime As Date, Optional strNotice As String = "*Nothing*") As Boolean
'------------------------------------------------
'���ܣ�����ԤԼ��Ϣ
'������ dtSchStartTime -- ԤԼ��ʼʱ��
'       dtSchEndTime -- ԤԼ����ʱ��
'       lngOrderID -- ҽ��ID
'       strName -- ��������
'       lngNumber -- ԤԼ���
'       lngSchDeviceID -- ԤԼ�豸ID
'       dtSegStartTime -- ��ʼʱ���
'       dtSegEndTime -- ����ʱ���
'       strNotice -- ���ע��
'���أ�True - �ɹ� �� False - ʧ��
'------------------------------------------------
    Dim strStartTime As String
    Dim strEndTime As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strNoticeSql As String
    
    On Error GoTo err
    
    'ÿ�α���ǰ���ȵ����Զ�ɾ�����ϸ�ԤԼ�ķ���
    strSQL = "Zl_Ӱ��ԤԼ�Զ�ɾ��"
    zlDatabase.ExecuteProcedure strSQL, "�Զ�ɾ�����ϸ��ԤԼ��¼"
    
    '�ȱ���ԤԼʱ���Ƿ�������������ظ�
    '��ʼʱ���һ���ӣ�����ʱ���һ���ӣ�
    strStartTime = Format(DateAdd("n", 1, dtSchStartTime), "yyyy-MM-dd hh:mm:ss")
    strEndTime = Format(DateAdd("n", -1, dtSchEndTime), "yyyy-MM-dd hh:mm:ss")
    
    strSQL = "Select ԤԼ�豸����,ԤԼ��ʼʱ��,ԤԼ����ʱ�� " _
        & " From Ӱ��ԤԼ��¼ Where ҽ��ID In " _
        & " (Select ID From ����ҽ����¼ Where ����ID =  " _
        & " (Select ����ID From ����ҽ����¼ Where ID = [1])) And ҽ��ID <> [1] And " _
        & " (ԤԼ��ʼʱ�� Between [2] And [3] or ԤԼ����ʱ�� Between [2] And [3] or " _
        & " [2] between ԤԼ��ʼʱ�� and ԤԼ����ʱ�� ) "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�������豸�ϵ�ԤԼ", lngOrderID, _
        CDate(strStartTime), CDate(strEndTime))

    If rsTemp.EOF = False Then
        If MsgBox("���� " & strName & " �� " & Format(rsTemp!ԤԼ��ʼʱ��, "hh:mm:ss") & " �� " & Format(rsTemp!ԤԼ����ʱ��, "hh:mm:ss") & " ���ʱ����Ѿ�ԤԼ�ˡ�" & nvl(rsTemp!ԤԼ�豸����) & "���ϵļ�飬" _
            & vbCrLf & vbCrLf & "�����޸ı��μ��ԤԼ��ʱ�䣬���ⷢ�����ʱ���ͻ��" & vbCrLf & vbCrLf & "�Ƿ�ȡ�����滼�� " & strName & " ��ԤԼʱ�䣿", vbYesNo, "���ԤԼ��ʾ") = vbYes Then
            Exit Function
        End If
    End If

    If strNotice = "*Nothing*" Then
        strNoticeSql = ""
    Else
        strNoticeSql = ",'" & strNotice & "'"
    End If
    
    strSQL = "Zl_Ӱ��ԤԼ��¼_����(" & lngOrderID & ",'" & lngNumber & "'," _
        & lngSchDeviceID & "," & zlStr.To_Date(dtSegStartTime) _
        & "," & zlStr.To_Date(dtSegEndTime) & "," & zlStr.To_Date(dtSchStartTime) _
        & "," & zlStr.To_Date(dtSchEndTime) & strNoticeSql & ")"
    zlDatabase.ExecuteProcedure strSQL, "������ԤԼ"
            
    funSaveSchedule = True
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub RefreshTimeProject(lngSchPlanID As Long)
'------------------------------------------------
'���ܣ�װ��ԤԼʱ���
'������lngSchPlanID -- ԤԼ����ID
'���أ���
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strStartTime As String
    Dim strEndTime As String
    Dim lngStartRow As Long
    Dim lngStartCol As Long
    Dim lngEndRow As Long
    Dim lngEndCol As Long
    Dim lngTimeProjectID As Long
    Dim iCount As Integer
    Dim blnChange As Boolean
    
    On Error GoTo err
    
    '���ԤԼ��ǩ
    Call unloadSchLabel
    
    mlngSchPlanID = lngSchPlanID
    
    ReDim Preserve mSchTimeProject(0) As Type_TimeProject
    iCount = 0
    mlngSchSum = 0
    
    strSQL = "select ID,��ʼʱ��,����ʱ��,ԤԼ����,���㷽�� from Ӱ��ԤԼʱ��ƻ� where ԤԼ����ID =[1] order by ��ʼʱ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����ԤԼʱ���", mlngSchPlanID)
    
    With vsfTime
        '����������ʱ�����ʾΪ��Ϣʱ��
        .Cell(flexcpBackColor, 1, 1, .Rows - 1, .Cols - 1) = mlngColorTabRest
        .Cell(flexcpData, 1, 1, .Rows - 1, .Cols - 1) = 0
        '�������ʾ����ʱ��
        Set mColordict = New Dictionary
        While rsTemp.EOF = False
            '��ȡʱ��
            strStartTime = Format(nvl(rsTemp!��ʼʱ��, Now), "hh:mm")
            strEndTime = Format(nvl(rsTemp!����ʱ��, Now), "hh:mm")
            lngTimeProjectID = rsTemp!ID
            
            '��ʱ��ƻ����浽������
            iCount = iCount + 1
            ReDim Preserve mSchTimeProject(iCount) As Type_TimeProject
            mSchTimeProject(iCount).lngID = rsTemp!ID
            mSchTimeProject(iCount).lngSchPlanID = mlngSchPlanID
            mSchTimeProject(iCount).dtStartTime = strStartTime
            mSchTimeProject(iCount).dtEndTime = strEndTime
            mSchTimeProject(iCount).lngSum = rsTemp!ԤԼ����
            mSchTimeProject(iCount).lngCalType = rsTemp!���㷽��
            mlngSchSum = mlngSchSum + mSchTimeProject(iCount).lngSum
            
            
            If strEndTime > strStartTime Then
                Call getRowColFromTime(strStartTime, True, lngStartRow, lngStartCol)
                Call getRowColFromTime(strEndTime, False, lngEndRow, lngEndCol)
                
                '��ʱ���ı�������Ҫ�ֳ�����������ʾһ��ʱ��ƻ�
                If (lngStartRow = lngEndRow) Then
                    '��һ��Сʱ֮�ڣ�ֻ��һ��
                    .Cell(flexcpBackColor, lngStartRow, lngStartCol, lngEndRow, lngEndCol) = mlngColorTabWork
                    .Cell(flexcpData, lngStartRow, lngStartCol, lngEndRow, lngEndCol) = lngTimeProjectID
                Else
                    '�Ȼ���һ��
                    .Cell(flexcpBackColor, lngStartRow, lngStartCol, lngStartRow, .Cols - 1) = mlngColorTabWork
                    .Cell(flexcpData, lngStartRow, lngStartCol, lngStartRow, .Cols - 1) = lngTimeProjectID '
                    If (lngEndRow - lngStartRow = 1) Then
                        'ֻ�����У����ڶ���
                        .Cell(flexcpBackColor, lngEndRow, 1, lngEndRow, lngEndCol) = mlngColorTabWork
                        .Cell(flexcpData, lngEndRow, 1, lngEndRow, lngEndCol) = lngTimeProjectID
                    Else
                        '���м��һ��
                        .Cell(flexcpBackColor, lngStartRow + 1, 1, lngEndRow - 1, .Cols - 1) = mlngColorTabWork
                        .Cell(flexcpData, lngStartRow + 1, 1, lngEndRow - 1, .Cols - 1) = lngTimeProjectID
                        '�����һ��
                        .Cell(flexcpBackColor, lngEndRow, 1, lngEndRow, lngEndCol) = mlngColorTabWork
                        .Cell(flexcpData, lngEndRow, 1, lngEndRow, lngEndCol) = lngTimeProjectID
                    End If
                End If
                
                If Not mColordict.Exists(lngTimeProjectID) Then
                    blnChange = Not blnChange
                    Call mColordict.Add(lngTimeProjectID, blnChange)
                End If

            Else
                '�������ʱ��С�ڿ�ʼʱ�䣬�򲻻����ʱ��ƻ�
            End If
            rsTemp.MoveNext
        Wend
    End With
    '����ԤԼ����û������ˢ��������ɫ
    Call ShowMouseTime(1, 1)
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Public Sub Init(lngUseType As Long)
'------------------------------------------------
'���ܣ����ⲿ���õĳ�ʼ���������ؼ��Ĳ���������Ҫ�����ݿ��ȡ���ݳ�ʼ����Ҫ�����е�ʱ�����
'������
'���أ���
'------------------------------------------------
    On Error GoTo err
    
    Call loadColor      '������ɫ
    Call LoadCalendar   '�����������ݼ�
    mlngUseType = lngUseType
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function RefreshSchedule(lngSchDeviceID As Long, dtDate As Date, lngOrderID As Long, _
    Optional lngSchPlanID As Long = 0, Optional blnMoved As Boolean = False) As Boolean
'------------------------------------------------
'���ܣ�ˢ��ԤԼʱ���
'������ lngSchDeviceID -- ԤԼ�豸ID
'       dtDate -- ԤԼʱ��
'       lngOrderID -- ҽ��ID
'       lngSchPlanID -- ��ѡ������0����ʾ��Ҫ���²�ѯ
'       blnMoved -- ��ѡ�������Ƿ��Ѿ�ת��
'���أ�True -- ���ԤԼģʽ�£��е�ǰҽ����ԤԼ��ǩ��False -- ���ԤԼģʽ�£��޵�ǰҽ����ԤԼ��ǩ��������ԤԼ����ģʽ
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim rsBaby As ADODB.Recordset
    Dim strFilter As String
    Dim iIndex As Integer
    Dim intPoolCount As Integer
    
    On Error GoTo err
    
    RefreshSchedule = False
    mlngSchDeviceID = lngSchDeviceID
    mlngOrderID = lngOrderID
    mdtSchDate = dtDate
    mblnMoved = blnMoved
    
    If lngSchDeviceID = 0 Then
        '����ʱ��ƻ����ʱ�䷽��
        Call RefreshTimeProject(0)
        Exit Function
    End If
    
    If lngSchPlanID = 0 Then
        mlngSchPlanID = GetSchPlanID(mlngSchDeviceID, dtDate, False, False)
    ElseIf lngSchPlanID = -1 Then
        lngSchPlanID = 0
        mlngSchPlanID = lngSchPlanID
    Else
        mlngSchPlanID = lngSchPlanID
    End If
    
    '����ʱ��ƻ����ʱ�䷽��
    Call RefreshTimeProject(mlngSchPlanID)
    
    '��ʾ�����ԤԼ���
    '��ѯԤԼ��¼
    strSQL = "Select d.ID, d.ҽ��ID, d.���, d.��������, d.ԤԼ��ʼʱ��, d.ԤԼ����ʱ��, " _
        & " d.ԤԼ��ʼʱ���, d.ԤԼ����ʱ���, b.����, b.ҽ������, b.Ӥ��, c.ִ�й��� " _
        & " From ����ҽ����¼ B, ����ҽ������ C,Ӱ��ԤԼ��¼ D Where b.id in " _
        & " (Select  a.ҽ��ID From Ӱ��ԤԼ��¼ A Where a.ԤԼ�豸ID = [1] And " _
        & " a.ԤԼ��ʼʱ�� Between [2] And [3] )And c.ҽ��id = b.id and d.ҽ��id=B.id Order By cast(d.��� as int)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�����ԤԼ��¼", mlngSchDeviceID, CDate(Format(dtDate, "yyyy-MM-dd 00:00:00")), CDate(Format(dtDate, "yyyy-MM-dd 23:59:59")))
        
    If rsTemp.EOF = False Then
        While rsTemp.EOF = False
            '��¼ԤԼ��Ϣ
            intPoolCount = UBound(mSchLabelPool) + 1
            ReDim Preserve mSchLabelPool(intPoolCount) As TYPE_SchLabel
            
            '��ȡӤ����Ϣ
            If rsTemp!Ӥ�� <> 0 Then
                strSQL = "Select A.����ʱ��,Nvl(B.Ӥ������, A.���� || '֮��' || Trim(To_Char(B.���, '9'))) As Ӥ������, B.Ӥ���Ա�, B.����ʱ��" & vbNewLine & _
                                 "  From ����ҽ����¼ A, ������������¼ B " & vbNewLine & _
                                 "  Where a.����ID = b.����ID  And b.��� = [2] And a.ID = [1]"
                Set rsBaby = zlDatabase.OpenSQLRecord(strSQL, "��ȡӤ����Ϣ", CLng(rsTemp!ҽ��ID), CLng(rsTemp!Ӥ��))
                mSchLabelPool(intPoolCount).str���� = "Ӥ����" & rsBaby!Ӥ������
            Else
                mSchLabelPool(intPoolCount).str���� = rsTemp!����
            End If
            mSchLabelPool(intPoolCount).lng��� = rsTemp!���
            mSchLabelPool(intPoolCount).lngҽ��ID = rsTemp!ҽ��ID
            mSchLabelPool(intPoolCount).dtStartTime = rsTemp!ԤԼ��ʼʱ��
            mSchLabelPool(intPoolCount).dtEndTime = rsTemp!ԤԼ����ʱ��
            mSchLabelPool(intPoolCount).strҽ������ = rsTemp!ҽ������
            mSchLabelPool(intPoolCount).dt��ʼʱ��� = rsTemp!ԤԼ��ʼʱ���
            mSchLabelPool(intPoolCount).dt����ʱ��� = rsTemp!ԤԼ����ʱ���
            mSchLabelPool(intPoolCount).lngBtnCount = 1
            mSchLabelPool(intPoolCount).isModified = False
            mSchLabelPool(intPoolCount).bln��ִ�� = IIf(nvl(rsTemp!ִ�й���, 0) = 0 Or nvl(rsTemp!ִ�й���, 0) = 1 Or nvl(rsTemp!ִ�й���, 0) = 2, False, True)
            mSchLabelPool(intPoolCount).bln�ѱ��� = True
            '����һ��ԤԼ��ǩ
            iIndex = CreateNewSchLabel()
            
            '�ڷ�ԤԼ��ǩ
            Call PutSchLabel(iIndex, intPoolCount)
            '
            '����Ǽ��ԤԼģʽ����Щ�����ݿ��ж�ȡ��ԤԼ��ǩ��ȫ������ֻ����ʾ
            If mlngUseType = Sch_UseType_���ԤԼ Then
                If mSchLabelPool(intPoolCount).lngҽ��ID <> mlngOrderID Then
                    Call setSchLabelEnable(iIndex, False)
                Else
                    '������������ҽ����ԤԼ��¼����¼ģ�����
                    mlngPoolIndex = intPoolCount
                    mlngBtnIndex = iIndex
                    'Call setSchLabelEnable(iIndex, IIf(mdtSchDate < Format(Now, "yyyy-mm-dd"), False, IIf(mSchLabelPool(intPoolCount).bln�ѱ��� = True, False, True)))
                    Call setSchLabelEnable(iIndex, IIf(mdtSchDate < Format(Now, "yyyy-mm-dd"), False, True))
                    Call setSchLabelToolTipText(iIndex)
                    RefreshSchedule = True  '���سɹ�
                End If
            ElseIf mlngUseType = Sch_UseType_ԤԼ���� Then
                Call ClearLocalParas
                Call setSchLabelEnable(iIndex, IIf(mdtSchDate < Format(Now, "yyyy-mm-dd"), False, True))
            End If
            rsTemp.MoveNext
        Wend
        If mlngUseType = Sch_UseType_���ԤԼ And RefreshSchedule = False Then
            Call ClearLocalParas
        End If
    Else
        'û�м�¼
        ReDim mSchLabelPool(0) As TYPE_SchLabel
        Call ClearLocalParas
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function NewSchedule(ByVal lngSchDeviceID As Long, ByRef dtSchDate As Date, _
    ByVal lngOrderID As Long, ByVal blnChangeDate As Boolean) As Boolean
'------------------------------------------------
'���ܣ�����ҽ��ID������һ���µļ��ԤԼ��ǩ�����ҽ���ǩ�Զ����ں��ʵ�λ��
'������ lngSchDeviceID -- ԤԼ�豸ID
'       dtSchDate -- ԤԼ����
'       lngOrderID -- ҽ��ID
'       blnChangeDate -- �Ƿ�ı�ԤԼ���ڣ�true--�����쿪ʼѰ�����ʺϵ�ԤԼ���ڣ�False--��������ԤԼ
'���أ�True -- �ɹ��� False -- ʧ��
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim rsBaby As ADODB.Recordset
    Dim rsSchedule As ADODB.Recordset
    Dim lngSchPlanID As Long
    Dim i As Integer
    Dim lngNumber As Long       '��С���õ�ԤԼ���
    Dim lngBaseNumber As Long   'ǰ�漸��ʱ��ο���ԤԼ�����ܺ�
    Dim dtLastEndTime As Date   '��һ��ԤԼ�Ľ���ʱ��
    Dim dtNextStartTime As Date '��һ��ԤԼ�Ŀ�ʼʱ��
    Dim dtStartTime As Date     '��ԤԼ�Ŀ�ʼʱ��
    Dim dtEndTime As Date       '��ԤԼ�Ľ���ʱ��
    Dim iIndex As Integer       '��ԤԼ��ǩ������
    Dim iPoolIndex As Integer   '��ԤԼ��ǩ��ԤԼ�ص�����
    Dim iTimeType As Integer    'ԤԼʱ�����㷽��
    Dim lngTimeLength As Long   'ԤԼʱ��
    Dim dtBeforeTime As Date    'ԤԼ���������ʱ�䣬����ǽ��죬����now+5����
    Dim blnFill As Boolean      '�Ƿ�����������ŵĿ�϶ʱ��
    Dim blnTimeOK As Boolean    '��������ʱ�����
    
    On Error GoTo err
    If Format(dtSchDate, "YYYY-MM-DD") < Format(Now, "YYYY-MM-DD") Then
        MsgBox "������ǰ�����ڣ������½�ԤԼ�� ", vbOKOnly, "���ԤԼ��ʾ"
        Exit Function
    End If
    blnFill = False
    
    If blnChangeDate = True Then
        '���Ҵ� ���� ��ʼ���󣬵�һ����ԤԼ����������
        lngSchPlanID = FindFirstSchDay(lngSchDeviceID, dtSchDate)
    Else
        lngSchPlanID = GetSchPlanID(lngSchDeviceID, dtSchDate, False, False)
    End If
    
    '�ҵ����ʵ�ԤԼ���ں�ˢ��ʱ���
    '����ˢ��ԤԼʱ�����ʾ��ǰ�Ѿ�ԤԼ�õ����
    If RefreshSchedule(lngSchDeviceID, dtSchDate, lngOrderID, IIf(lngSchPlanID = 0, -1, lngSchPlanID)) = True Then
        Exit Function   '����������Ѿ���ԤԼ���Ͳ�����ʾ�µ�ԤԼ��ǩ
    End If
    
    If lngSchPlanID = 0 Then
        Exit Function
    End If
    
    '������ţ��ҵ����ʺϵ�ԤԼλ��
    If UBound(mSchTimeProject) = 0 Then
        MsgBox "�޷�ԤԼ����ԤԼ����û��ʱ��ƻ����������ú�ԤԼʱ��ƻ���ԤԼ����ID=" & lngSchPlanID, vbOKOnly, "���ԤԼ��ʾ"
        Exit Function
    End If
    
    If Format(dtSchDate, "YYYY-MM-DD") < Format(Now, "YYYY-MM-DD") Then
        '������ʾ�ƶ�����ǰ��
        'MsgBox "������ǰ�����ڣ������½�ԤԼ�� ", vbOKOnly, "���ԤԼ��ʾ"
        'Exit Function
    ElseIf Format(dtSchDate, "YYYY-MM-DD") = Format(Now, "YYYY-MM-DD") Then
        dtBeforeTime = DateAdd("n", 5, Now)
    Else
        dtBeforeTime = Format(dtSchDate, "YYYY-MM-DD") & " 00:00:00"
    End If
    
    '����ʱ��˳���������ʱ��ƻ������ҿ���ԤԼ������ʱ��
    '������С�Ŀ�ԤԼ���
    strSQL = "select ���,ԤԼ��ʼʱ��,ԤԼ����ʱ�� from Ӱ��ԤԼ��¼ where ԤԼ�豸ID=[1] and ԤԼ��ʼʱ�� between [2] and [3] order by cast(��� as int)"
    Set rsSchedule = zlDatabase.OpenSQLRecord(strSQL, "������С�Ŀ�ԤԼ���", mlngSchDeviceID, _
        CDate(Format(dtSchDate, "yyyy-MM-dd") & " 00:00:00"), CDate(Format(dtSchDate, "yyyy-MM-dd") & " 23:59:59"))
    
    While blnTimeOK = False
        dtStartTime = 0
        dtNextStartTime = Format(dtSchDate, "yyyy-MM-dd") & " 23:59:59"
        iTimeType = 0
        lngTimeLength = 0
        If Not rsSchedule.EOF Then rsSchedule.MoveFirst
        lngNumber = 1
        
        Do While rsSchedule.EOF = False
            If rsSchedule!��� = lngNumber Then
                dtLastEndTime = rsSchedule!ԤԼ����ʱ��
                lngNumber = lngNumber + 1       '�������������ţ�����պò鵽��һ���պ�
            ElseIf rsSchedule!ԤԼ��ʼʱ�� < dtBeforeTime Then
                'ʱ��̫�磬����֮�󣬼��������
                dtLastEndTime = rsSchedule!ԤԼ����ʱ��
                lngNumber = rsSchedule!��� + 1
            Else
                '�����ǽ����ŵ���һ�����룬��������¼����
'                If rsSchedule!��� = lngNumber + 1 Then
'                    blnFill = True
'                End If
                dtNextStartTime = rsSchedule!ԤԼ��ʼʱ��
                Exit Do
            End If
            rsSchedule.MoveNext
        Loop
            
        If dtLastEndTime < dtBeforeTime Then dtLastEndTime = dtBeforeTime
        
        '������С��ţ�ԤԼ��ʼʱ�䣬���������ʱ���
        i = 1
        lngBaseNumber = 0
        While i <= UBound(mSchTimeProject) And dtStartTime = 0
            If (lngNumber > mSchTimeProject(i).lngSum + lngBaseNumber) Or _
                (Format(dtBeforeTime, "HH:MM") > Format(mSchTimeProject(i).dtEndTime, "HH:MM")) Then
                lngBaseNumber = lngBaseNumber + mSchTimeProject(i).lngSum
                i = i + 1
            Else
                '�ҵ����ʱ��Σ�����ѭ��
                If (Format(dtSchDate, "YYYY-MM-DD") & " " & Format(mSchTimeProject(i).dtStartTime, "HH:MM:SS") > dtLastEndTime) Then
                    dtStartTime = Format(dtSchDate, "YYYY-MM-DD") & " " & Format(mSchTimeProject(i).dtStartTime, "HH:MM:SS")
                Else
                    dtStartTime = dtLastEndTime
                End If
                iTimeType = mSchTimeProject(i).lngCalType
                lngTimeLength = DateDiff("n", mSchTimeProject(i).dtStartTime, mSchTimeProject(i).dtEndTime) / mSchTimeProject(i).lngSum
'                If blnFill = True Then
'                    If Format(dtNextStartTime, "HH:MM") < mSchTimeProject(i).dtStartTime _
'                        Or mSchTimeProject(i).dtEndTime < Format(dtNextStartTime, "HH:MM") Then
'                        blnFill = False
'                    End If
'                End If
            End If
        Wend
        
        '���ʱ�䳬����ʱ��ƻ�����ʾ�û��Ƿ����ƻ���ԤԼ
        If dtStartTime = 0 Then
            '���һ��ʱ����Ƿ������
            If Format(dtBeforeTime, "HH:MM") < Format(mSchTimeProject(UBound(mSchTimeProject)).dtEndTime, "HH:MM") Then
                dtStartTime = dtLastEndTime
                iTimeType = mSchTimeProject(UBound(mSchTimeProject)).lngCalType
                lngTimeLength = DateDiff("n", mSchTimeProject(UBound(mSchTimeProject)).dtStartTime, mSchTimeProject(UBound(mSchTimeProject)).dtEndTime) / mSchTimeProject(UBound(mSchTimeProject)).lngSum
            Else
                If MsgBox("�����Ѿ�û�п���ԤԼʱ��ƻ����Ƿ��ڼƻ���ԤԼ��", vbYesNo, "���ԤԼ��ʾ") = vbNo Then
                    Exit Function
                End If
                dtStartTime = dtBeforeTime
            End If
        End If
        
        '����ԤԼ�Ľ���ʱ��
        '��������ŵģ��պ����������м䣬��ֱ������������������ʱ���϶
'        If blnFill = True Then
'            dtEndTime = dtNextStartTime
'        Else
            '�����ݿ⣬��ȡ����ʱ�䣬���˴�ƽ������ǰ���Ѿ��������
            If lngTimeLength <> 0 And iTimeType = 1 Then
                dtEndTime = DateAdd("n", lngTimeLength, dtStartTime)
            Else
                '����Ŀ�ۼ�
                strSQL = "select b.���ʱ�� from ����ҽ����¼ a ,Ӱ��ԤԼ��Ŀ b where a.������Ŀid = b.������Ŀid and a.id =[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯԤԼ��Ŀ���ʱ��", lngOrderID)
                
                If rsTemp.EOF = False Then
                    dtEndTime = DateAdd("n", rsTemp!���ʱ��, dtStartTime)
                End If
            End If
'        End If
        
        '���ʱ��ƻ��ĳ��Ȳ�����������Ѱ����һ��λ�ú����
        If dtEndTime > dtNextStartTime Then
            dtBeforeTime = dtEndTime
        Else
            blnTimeOK = True
        End If
    Wend
    
    '�����ݿ��в�ѯҽ����Ϣ
    strSQL = "select id,����,ҽ������,Ӥ��  from ����ҽ����¼  where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯҽ������", lngOrderID)
    
    '����ԤԼ����������ԤԼ��Ϣ��¼��ԤԼ����
    iPoolIndex = UBound(mSchLabelPool) + 1
    ReDim Preserve mSchLabelPool(iPoolIndex) As TYPE_SchLabel
    
    '��ȡӤ����Ϣ
    If rsTemp!Ӥ�� <> 0 Then
        strSQL = "Select A.����ʱ��,Nvl(B.Ӥ������, A.���� || '֮��' || Trim(To_Char(B.���, '9'))) As Ӥ������, B.Ӥ���Ա�, B.����ʱ��" & vbNewLine & _
                    "  From ����ҽ����¼ A, ������������¼ B " & vbNewLine & _
                    "  Where a.����ID = b.����ID  And b.��� = [2] And a.ID = [1]"
        Set rsBaby = zlDatabase.OpenSQLRecord(strSQL, "��ȡӤ����Ϣ", lngOrderID, CLng(rsTemp!Ӥ��))
        mSchLabelPool(iPoolIndex).str���� = "Ӥ����" & rsBaby!Ӥ������
    Else
        mSchLabelPool(iPoolIndex).str���� = rsTemp!����
    End If
    mSchLabelPool(iPoolIndex).lng��� = lngNumber
    mSchLabelPool(iPoolIndex).lngҽ��ID = rsTemp!ID
    mSchLabelPool(iPoolIndex).dtStartTime = dtStartTime
    mSchLabelPool(iPoolIndex).dtEndTime = dtEndTime
    mSchLabelPool(iPoolIndex).strҽ������ = rsTemp!ҽ������
    mSchLabelPool(iPoolIndex).lngBtnCount = 1
    mSchLabelPool(iPoolIndex).bln��ִ�� = False
    mSchLabelPool(iPoolIndex).bln�ѱ��� = False
    
    '���µ���ԤԼ��˳��
    iPoolIndex = ResortSchLabelPool(iPoolIndex)
    
    '����һ��ԤԼ��ǩ���ڷ���dtStartTimeλ��
    iIndex = CreateNewSchLabel()
    
    '�ڷ�ԤԼ��ǩ
    Call PutSchLabel(iIndex, iPoolIndex)
    
    '��¼ģ�����
    mlngPoolIndex = btnSchLabel(iIndex).tag
    mlngBtnIndex = iIndex
    
    Call setSchLabelToolTipText(iIndex)
    
    NewSchedule = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'-------------------------------------------------------------
'   �ڲ�˽�з���
'-------------------------------------------------------------

Private Sub btnSchLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then mlngTimeProjectID = 0
    If Button = 1 And btnSchLabel(Index).Visible = True Then
        mlngOriginLeft = btnSchLabel(mSchLabelPool(btnSchLabel(Index).tag).lngBtnIndex).Left
        mlngOriginTop = btnSchLabel(mSchLabelPool(btnSchLabel(Index).tag).lngBtnIndex).Top
        mlngOriginMinute = DateDiff("n", mSchLabelPool(btnSchLabel(Index).tag).dtStartTime, mSchLabelPool(btnSchLabel(Index).tag).dtEndTime)
        mlngOriginWidth = btnSchLabel(Index).Width
        mlngPoolIndex = btnSchLabel(Index).tag
        
        mlngBaseX = X
        mlngBaseY = Y
        mlngDownX = X
        
        Call setSchLabelSelectTag(Index)
        
        '�ѱ�ǩ��ʾ����ǰ��
        Call setSchLabelZorder(Index)
        '���ҽ��ID�����ı䣬�򴥷�OnChangeOrder�¼�
        If mlngOrderID <> mSchLabelPool(btnSchLabel(Index).tag).lngҽ��ID Then
            RaiseEvent OnChangeOrder(mSchLabelPool(btnSchLabel(Index).tag).lngҽ��ID, btnSchLabel(Index).ToolTipText)
        End If
        mlngOrderID = mSchLabelPool(btnSchLabel(Index).tag).lngҽ��ID
        
        If btnSchLabel(Index).MousePointer = vbSizeWE Then
            '����ԤԼ��ǩ���
            If CanResizeLabel(Index) = True Then
                mblnSizingLabel = True
            Else
                mblnSizingLabel = False
            End If
        Else
            '��קԤԼ��ǩ
            mblnRestorePos = False
            mblnDragingLabel = True
        End If
    End If
End Sub

Private Sub btnSchLabel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngWidth As Long
    Dim iPoolIndex As Long
    
    

    On Error GoTo err
    If mblnDragingLabel = True And Button = 1 Then
        If mlngBaseX <> X Or mlngBaseY <> Y Then
            Call MoveBtnLabels(Index, mlngBaseX, mlngBaseY, IIf(X < 0, 0, X), Y)
        End If
    ElseIf mblnSizingLabel = True And Button = 1 Then
        '���ڵ���ԤԼ��ǩ�Ŀ��
        
        lngWidth = mlngOriginWidth + (X - mlngBaseX)
        If lngWidth > 0 And lngWidth < (vsfTime.Width - vsfTime.ColWidth(0)) Then
            btnSchLabel(Index).Width = lngWidth
        End If
    Else
        '������ֻ�Ǿ���ԤԼ��ǩ���������ڱ�ǩ���ұߣ���ʾ���ҵ��������ָ��
        If X > btnSchLabel(Index).Width - 100 Then
            btnSchLabel(Index).MousePointer = vbSizeWE
        Else
            btnSchLabel(Index).MousePointer = vbDefault
        End If
        
    On Error Resume Next
    
    '���¼��ر�ǩ��Ӧ����ʾ��Ϣ
    Call setSchLabelToolTipText(Index)
    
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub btnSchLabel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngFirstIndex As Long
    lngFirstIndex = mSchLabelPool(btnSchLabel(Index).tag).lngBtnIndex
    
    If Button = 1 Then mlngTimeProjectID = 0
    Call SetColor(0, True)
    If mblnDragingLabel = True Then
        
'        '���ڱ�ǩ�飬���� mlngBaseX��ÿһ���ƶ��󣬶����������������ʹ��mlngDownX���жϱ�ǩ�Ƿ���ı��ƶ���
'        If mlngDownX <> X Or mlngBaseY <> Y Then
'            '������ǩλ��
'            Call AdjustLabelPos(Index)
'        End If
'
'        mlngPoolIndex = btnSchLabel(lngFirstIndex).tag
        mblnDragingLabel = False
'        mSchLabelPool(btnSchLabel(Index).tag).bln�ѱ��� = False
'        Call setSchLabelToolTipText(lngFirstIndex)
'        RaiseEvent OnSchLabelModifed(lngFirstIndex)
'
'        If Button = 1 Then
'            If mlngBaseX <> X Or mlngBaseY <> Y Then Call MoveBtnLabelsAuto(Index, mlngBaseX, mlngBaseY, IIf(X < 0, 0, X), Y)
'        End If
        
        Call SetMouseTimePro(0, 0, True)
    ElseIf mblnSizingLabel = True Then
        '��������ԤԼ��ǩ�Ŀ��
        mblnSizingLabel = False
        
        '���µ�ԤԼʱ����¼��ԤԼ����
        '�����ǵ���������ǩ�����Ǳ�ǩ�飬����ֻ�ܵ�����ǩ���ұ߽磬�������λ�ÿ϶��ǽ���ʱ��
        mSchLabelPool(btnSchLabel(Index).tag).dtEndTime = Format(mSchLabelPool(btnSchLabel(Index).tag).dtEndTime, "YYYY-MM-DD") & " " & Format(GetTimeFromXY(btnSchLabel(Index).Left + btnSchLabel(Index).Width, btnSchLabel(Index).Top), "HH:MM")
            
        If IsLabelOverlap(Index) = True Then
            '������ʾ��ֱ�ӽ���ǩ�ƶ���ԭ����λ��
            Call RestoreLabelPos(Index)
        End If
        
        mSchLabelPool(btnSchLabel(Index).tag).bln�ѱ��� = False
        Call setSchLabelToolTipText(lngFirstIndex)
        RaiseEvent OnSchLabelModifed(lngFirstIndex)
    ElseIf Button = 2 And (mlngUseType = Sch_UseType_���ԤԼ Or mlngUseType = Sch_UseType_ԤԼ����) Then
        '�����Ҽ��˵�
        If mlngUseType = Sch_UseType_���ԤԼ Then
            menuPopupScheduleModi.Visible = False
        Else
            menuPopupScheduleModi.Visible = True
        End If
        Call PopupMenu(menuSchedulePopup)
    End If

End Sub

Private Sub loadTimeTable()
'------------------------------------------------
'���ܣ�װ��ʱ���ı���ʽ�ͻ�������
'������
'���أ���
'------------------------------------------------
    Dim i As Integer
    
    With vsfTime
        .Rows = 25
        .Cols = 13
        .FixedRows = 1
        .FixedCols = 1
        .AllowUserResizing = flexResizeNone
        .Editable = flexEDNone
        .ScrollBars = flexScrollBarNone
        .CellAlignment = flexAlignLeftCenter
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .AllowSelection = False
        
        For i = 0 To 23
            .TextMatrix(i + 1, 0) = IIf(i < 10, "0" & i, i) & ":00"
        Next i
        
        For i = 0 To 11
            .TextMatrix(0, i + 1) = IIf(i * 5 < 10, "0" & i * 5, i * 5)
        Next i
        
    End With
End Sub

Private Sub ResizeTimeTable()
'------------------------------------------------
'���ܣ�����ʱ����λ��
'������
'���أ���
'------------------------------------------------
    Dim lngTemp  As Long
    Dim i As Integer
    
    On Error GoTo err
    
    '�ȵ���ʱ����λ��
    With vsfTime
        .Left = ScaleLeft
        .Top = ScaleTop
        .Height = ScaleHeight
        .Width = ScaleWidth - 20
        .RowHeight(0) = .Height / 25
        If .RowHeight(0) < 400 Then .RowHeight(0) = 400
        lngTemp = (.Height - .RowHeight(0) - 100) / 24
        For i = 0 To 23
            .RowHeight(i + 1) = lngTemp
        Next i
        
        .RowHeight(0) = .Height - (.RowHeight(1) * 24) - 80
        
        '���еĿ��
        .ColWidth(0) = .Width / 13
        If .ColWidth(0) < 600 Then .ColWidth(0) = 600
        lngTemp = (.Width - .ColWidth(0)) / 12
        For i = 0 To 11
            .ColWidth(i + 1) = lngTemp
        Next i
        .ColWidth(0) = .Width - (.ColWidth(1) * 12)
    End With
    
    '�ٵ���ԤԼ��ǩ��λ��
    For i = 1 To UBound(mSchLabelPool)
        Call PutSchLabel(mSchLabelPool(i).lngBtnIndex, i)
    Next i
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub timColor_Timer()
'ˢ�µ�ǰ�ƶ�����ʱ���
    '�����б��г�����ɫѡ�п�
    vsfTime.Row = 0
    vsfTime.Col = 0
    If mblnDragingLabel Then
        Call ShowNowSchTimeProject
    Else
        If mlngUseType = Sch_UseType_���ԤԼ Then
            Call ShowMouseTime(mlngMouseX, mlngMouseY)
        End If
    End If
End Sub

Private Sub ShowNowSchTimeProject()
'չ�ֵ�ǰ������ڵ�ʱ���
    On Error GoTo errH
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngTimeProjectID As Long
    Dim i As Integer
    Dim intTMP As Integer, iFirstIndex As Integer, iPoolIndex As Integer
    
    intTMP = -1
    
    On Error Resume Next
    For i = 0 To btnSchLabel.Count - 1
        If btnSchLabel(i).ToolTipText = mstrOrderInfo Then
            intTMP = i
            Exit For
        End If
    Next
    
    On Error GoTo errH
    If intTMP = -1 Then Exit Sub
    If btnSchLabel(intTMP).HelpContextID <> 0 Then   '�Ǳ�ǩ��
        '�����ҵ������ǩ���еĵ�һ����ǩ����
        iFirstIndex = mSchLabelPool(btnSchLabel(intTMP).tag).lngBtnIndex
    Else    '���Ǳ�ǩ�飬��ǰ�������ǵ�һ������
        iFirstIndex = intTMP
    End If
    iPoolIndex = btnSchLabel(iFirstIndex).tag
    
    lngRow = GetRowsFromY(btnSchLabel(intTMP).Top)
    lngCol = GetColsFromX(btnSchLabel(intTMP).Left)
    lngTimeProjectID = vsfTime.Cell(flexcpData, lngRow, lngCol)
    If lngTimeProjectID = 0 Then
        Call SetColor(lngTimeProjectID, True)
    Else
        mSchLabelPool(iPoolIndex).lngTimeProjectID = lngTimeProjectID
        If mlngTimeProjectID <> lngTimeProjectID Then
            Call SetColor(lngTimeProjectID, False)
        End If
    End If
    
    Exit Sub
errH:
    If InStr(err.Description, "�ؼ�����Ԫ��") > 0 Then
        Resume Next
    Else
        Call err.Raise(err.Number, , err.Description)
        Resume
    End If
End Sub

Private Sub UserControl_Initialize()
    Call InitSchedule
End Sub

Private Sub UserControl_Resize()
    Call ResizeTimeTable
End Sub

Private Sub UserControl_Terminate()
    Set mrsCalendar = Nothing
    Set mColordict = Nothing
End Sub

Private Sub vsfTime_DragDrop(Source As Control, X As Single, Y As Single)
    
    Dim lngTop As Long
    
    On Error GoTo err
    
    If Source.Name = "btnSchLabel" Then
        Source.Left = X - Source.Width / 2
        
        '����ԤԼ��ǩ�ڷ���ʱ����еĸ߶ȣ���Ҫ��ʱ�����и�ƽ��
        lngTop = AdjustSchLabelTop(Y - Source.Height / 2)
        
        Source.Top = lngTop
        Source.Visible = True
    End If
    mblnDragingLabel = False
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function AdjustSchLabelTop(ByVal lngTop As Long) As Long
'------------------------------------------------
'���ܣ����ݵ�ǰ��λ�ã�΢��ԤԼ��ǩ��TOP���ѱ�ǩ�ڷ���ʱ����ĳһ����
'������ lngTop --- ��ǰ������ڵ�Yλ��
'���أ���ʱ����ڣ���ӽ����λ�õ����׵�Yֵ
'------------------------------------------------
    Dim i As Integer
    
    On Error GoTo err
        
    '����������У����Զ��ŵ���һ��
    If lngTop < vsfTime.RowHeight(0) Or lngTop > vsfTime.Height Then
        AdjustSchLabelTop = mlngOriginTop
        mblnRestorePos = True
    Else
        For i = 0 To 22
            If lngTop < vsfTime.RowHeight(0) + vsfTime.RowHeight(1) * i Then
                Exit For
            End If
        Next i
        If Abs(lngTop - (vsfTime.RowHeight(0) + vsfTime.RowHeight(1) * (i - 1))) > _
            Abs(lngTop - (vsfTime.RowHeight(0) + vsfTime.RowHeight(1) * i)) Then
            AdjustSchLabelTop = vsfTime.RowHeight(0) + vsfTime.RowHeight(1) * i
        Else
            AdjustSchLabelTop = vsfTime.RowHeight(0) + vsfTime.RowHeight(1) * (i - 1)
        End If
        AdjustSchLabelTop = AdjustSchLabelTop + 30
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function AdjustSchLabelLeft(ByVal lngLeft As Long) As Long
'------------------------------------------------
'���ܣ����ݵ�ǰ��λ�ã�΢��ԤԼ��ǩ��Left��ȷ����ǩ�����Ƴ�ʱ���Χ
'������ lngLeft --- ��ǰ������ڵ�Xλ��
'���أ���ʱ����ڣ������λ�õ�Xֵ
'------------------------------------------------
    Dim i As Integer
    
    On Error GoTo err
        
    '����������У���Ż�ԭ����λ��
    If lngLeft <= vsfTime.ColWidth(0) Or lngLeft >= vsfTime.Width Then
        AdjustSchLabelLeft = mlngOriginLeft
        mblnRestorePos = True
    Else
        AdjustSchLabelLeft = lngLeft
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function getRowColFromTime(ByVal strTime As String, ByVal blnStart As Boolean, _
    ByRef lngRow As Long, ByRef lngCol As Long) As Boolean
'------------------------------------------------
'���ܣ�����ʱ�䣬�����Ӧ�ĸ����к���λ��
'������ strTime -- �����ʱ�䣬��ʽΪ��hh:mm��
'       blnStart -- �Ƿ�ʼʱ�� true--��ʼʱ�䣻false--����ʱ��
'       lngRow -- ��ʱ����е��к�
'       lngCol -- ��ʱ����е��к�
'���أ���
'------------------------------------------------
    Dim lngHour As Long
    Dim lngMinute As Long
    
    On Error GoTo err
    
    If UBound(Split(strTime, ":")) <> 1 Then
        getRowColFromTime = False
        Exit Function
    End If
    
    lngHour = Split(strTime, ":")(0)
    lngMinute = Split(strTime, ":")(1)
    
    '����ǽ���ʱ�䣬��ʱ��ķ��ӵ���0����������ǰһ��Сʱ����
    If blnStart = False And lngMinute = 0 Then
        lngRow = lngHour
        lngCol = 12
    Else
        lngRow = lngHour + 1
        lngCol = Int(lngMinute / 5 + 1) 'ȷ��5�ı�����������һ��
    End If
    
    If blnStart = False And (lngMinute Mod 5 = 0) And (lngMinute <> 0) Then
        '����ǽ���ʱ�䣬������5������������Ҫ��ǰ��һ��
        lngCol = lngCol - 1
    End If
    
    getRowColFromTime = True
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub vsfTime_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mlngUseType = Sch_UseType_���ԤԼ Then
        If mblnDragingLabel Then
            mlngMouseX = 0
            mlngMouseY = 0
        Else
            mlngMouseX = X
            mlngMouseY = Y
        End If
    End If
End Sub

Private Sub vsfTime_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngTimeProjectID As Long
    
    If Button = 2 And mlngUseType = Sch_UseType_ԤԼ���� And mlngSchPlanID <> 0 Then       '����Ҽ������������Ҽ��˵�
        '�ȿ��Ʋ˵��Ŀɼ���
        
        With vsfTime
            mlngCol = GetColsFromX(X)
            mlngRow = GetRowsFromY(Y)
            If mlngCol >= 1 And mlngRow >= 1 Then
                lngTimeProjectID = .Cell(flexcpData, mlngRow, mlngCol, mlngRow, mlngCol)
                If lngTimeProjectID = 0 Then
                    menuPopupTimeProjectModi.Visible = False
                    menuPopupTimeProjectDel.Visible = False
                Else
                    menuPopupTimeProjectModi.Visible = True
                    menuPopupTimeProjectDel.Visible = True
                End If
            End If
        End With
    
        Call PopupMenu(menuTimePopup)
    End If
    
    If Button = 1 Then
        If mlngUseType = Sch_UseType_���ԤԼ Then
            Call SetMouseTimePro(CLng(X), CLng(Y), False)
        End If
    End If

End Sub

Private Sub menuPopupTimeProjectAdd_Click()
    On Error Resume Next
    
    RaiseEvent OnMenuTimeProjectAdd
    
    Call RefreshTimeProject(mlngSchPlanID)
    err.Clear
End Sub

Private Sub menuPopupTimeProjectModi_Click()
    Dim lngTimeProjectID As Long
    
    On Error Resume Next
    
    '��ʱ����У���ȡʱ��ƻ�ID
    If mlngCol >= 1 And mlngRow >= 1 Then
        lngTimeProjectID = vsfTime.Cell(flexcpData, mlngRow, mlngCol, mlngRow, mlngCol)
        RaiseEvent OnMenuTimeProjectModify(lngTimeProjectID)
    End If
    Call RefreshTimeProject(mlngSchPlanID)
    err.Clear
End Sub

Private Sub menuPopupTimeProjectDel_Click()
    'ɾ��ʱ��ƻ�
    Dim strSQL As String
    Dim lngTimeProjectID As Long
    Dim blnCancel As Boolean
    
    On Error Resume Next
    
    RaiseEvent OnMenuTimeProjectBeforeDel(blnCancel)
    
    If blnCancel = True Then
        Exit Sub
    End If
    
    '��ʱ����У���ȡʱ��ƻ�ID
    If mlngCol >= 1 And mlngRow >= 1 Then
        lngTimeProjectID = vsfTime.Cell(flexcpData, mlngRow, mlngCol, mlngRow, mlngCol)
        strSQL = "Zl_Ӱ��ԤԼʱ��ƻ�_ɾ��(" & lngTimeProjectID & ")"
        zlDatabase.ExecuteProcedure strSQL, "ɾ��ʱ��ƻ�"
        Call RefreshTimeProject(mlngSchPlanID)
    End If
    
    err.Clear
End Sub

Private Sub menuPopupTimeProjectColor_Click()
    On Error Resume Next
    
    RaiseEvent OnMenuTimeProjectSetColor
    Call loadColor
    
    err.Clear
End Sub

Private Sub menuPopupScheduleModi_Click()
    On Error Resume Next
    
    RaiseEvent OnMenuScheduleModify
    
    err.Clear
End Sub

Private Sub menuPopupSchedulePrint_Click()
    On Error Resume Next
    
    RaiseEvent OnMenuSchedulePrint
    
    err.Clear
End Sub

Private Sub InitSchedule()
'------------------------------------------------
'���ܣ���ʼ��ԤԼʱ���ؼ�
'������
'���أ���
'------------------------------------------------
    On Error GoTo err
    
    btnSchLabel(0).Width = 0
    btnSchLabel(0).Height = 0
    btnSchLabel(0).Visible = False
    
    Call loadTimeTable
    mlngSchLabelIndex = 0
    mIsReadOnly = False
    ReDim Preserve mSchLabelPool(0) As TYPE_SchLabel
    mstrModifiedOrderID = ""
    mlngFontSize = btnSchLabel(0).Font.Size
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub loadColor()
'------------------------------------------------
'���ܣ������ݿ��ȡʱ�����ɫ����
'������
'���أ���
'------------------------------------------------
    
    On Error GoTo err
    
    '�����ݿ��ж�ȡ���ù�����ɫ
    mlngColorTabWork = zlDatabase.GetPara("���ԤԼʱ�����ʱ����ɫ", glngSys, 1292, "8421376")
    mlngColorTabRest = zlDatabase.GetPara("���ԤԼʱ�����Ϣʱ����ɫ", glngSys, 1292, "16777215")
    mlngColorLblWaiting = zlDatabase.GetPara("���ԤԼ��ǩ��ԤԼ��ɫ", glngSys, 1292, "0")
    mlngColorLblDone = zlDatabase.GetPara("���ԤԼ��ǩ�������ɫ", glngSys, 1292, "12632256")
    mlngColorLblPassed = zlDatabase.GetPara("���ԤԼ��ǩ�ѹ�����ɫ", glngSys, 1292, "255")
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub unloadSchLabel()
'------------------------------------------------
'���ܣ����ԤԼ��ǩ��ԤԼ��ǩ�������ǲ�������
'������
'���أ���
'------------------------------------------------
    Dim i As Integer
    Dim btnLabel As CommandButton
    
    On Error GoTo err
    
    'ж��ԤԼ��ǩ
    For Each btnLabel In btnSchLabel
        If btnLabel.Index <> 0 Then
            Call Unload(btnLabel)
        End If
    Next
    
    '�Ѽ���������
    mlngSchLabelIndex = btnSchLabel.Count - 1
    
    '���ԤԼ��
    ReDim Preserve mSchLabelPool(0) As TYPE_SchLabel
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub PutSchLabel(ByVal iBtnIndex As Integer, ByVal iPoolIndex As Integer)
'------------------------------------------------
'���ܣ���������Ŀ�ʼʱ��ͽ���ʱ�䣬�ڷ�ԤԼ��ǩ��Ȼ������ŵ���ԤԼ�ص�˳��
'������ iBtnIndex �� ԤԼ��ǩ������
'       iPoolIndex -- ԤԼ�ص�����
'���أ���
'------------------------------------------------
    Dim strStartTime As String
    Dim strEndTime As String
    Dim intSHour As Integer
    Dim intSMinute As Integer
    Dim intEHour As Integer
    Dim intEMinute As Integer
    Dim lngSX As Long
    Dim lngSY As Long
    Dim lngEX As Long
    Dim lngEY As Long
    Dim iNewIndex As Integer
    Dim iPreIndex As Integer
    Dim i As Integer
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngTimeProjectID As Long        'ʱ��ƻ�ID
    Dim lngColor As Long
    
    
    On Error GoTo err
    '����ԤԼ�غ�ԤԼ��ǩ������
    mSchLabelPool(iPoolIndex).lngBtnIndex = iBtnIndex
        
    btnSchLabel(iBtnIndex).Caption = mSchLabelPool(iPoolIndex).str����
    btnSchLabel(iBtnIndex).tag = iPoolIndex
    
    '��ԤԼ�ػ�ȡ��ʼʱ��ͽ���ʱ��
    strStartTime = Format(mSchLabelPool(iPoolIndex).dtStartTime, "HH:MM")
    strEndTime = Format(mSchLabelPool(iPoolIndex).dtEndTime, "HH:MM")
    
    '������ɫ
    If mSchLabelPool(iPoolIndex).bln��ִ�� = True Then
        lngColor = mlngColorLblDone
    ElseIf Format(mdtSchDate, "YYYY-MM-DD") & " " & Format(strStartTime, "HH:MM:SS") < Format(Now, "YYYY-MM-DD HH:MM:SS") Then
        lngColor = mlngColorLblPassed
    Else
        lngColor = mlngColorLblWaiting
    End If
    
    '���ݿ�ʼʱ�䣬�����ǩ�Ŀ�ʼλ��
    intSHour = Val(Left(strStartTime, 2))
    intSMinute = Val(Mid(strStartTime, 4))
    intEHour = Val(Left(strEndTime, 2))
    intEMinute = Val(Mid(strEndTime, 4))
    
    'ʹ��ͳһ�ķ�������ʱ����ȡX,Y��λ�ã�����ʱ���Ƿ���Ҫר�Ŵ������������һ�У�Ӧ����ǰ����һ�н���
    Call GetXYFromTime(strStartTime, lngSX, lngSY)
    Call GetXYFromTime(strEndTime, lngEX, lngEY)
        
    '��ǩ�Ƿ���ͬһ����ʾ
    If (intEHour = intSHour) Or (intEHour - intSHour = 1 And intEMinute = 0) Then '��ǩֻ��һ��,�����жϿ�ʼ�ͽ���ʱ�䣬�Ƿ�Ϊͬһ��Сʱ

        btnSchLabel(iBtnIndex).Left = lngSX
        btnSchLabel(iBtnIndex).Top = lngSY
        If intEMinute = 0 Then
            btnSchLabel(iBtnIndex).Width = vsfTime.Width - lngSX
        Else
            btnSchLabel(iBtnIndex).Width = lngEX - lngSX
        End If
        
        btnSchLabel(iBtnIndex).Height = vsfTime.RowHeight(1)
        '���ԭ���Ǹ���ǩ�飬��ж�ض���ı�ǩ
        If btnSchLabel(iBtnIndex).HelpContextID <> 0 Then
            Call DelUnUseLabel(iBtnIndex, btnSchLabel(iBtnIndex).HelpContextID)
        End If
        btnSchLabel(iBtnIndex).HelpContextID = 0      '���õ��б�ǩ�ı��
        btnSchLabel(iBtnIndex).BackColor = lngColor
    Else    '��ǩ�ж���
        '����һ����ǩ��
        '�Ȱڷŵ�һ�б�ǩ
        btnSchLabel(iBtnIndex).Left = lngSX
        btnSchLabel(iBtnIndex).Top = lngSY
        btnSchLabel(iBtnIndex).Width = vsfTime.Width - lngSX
        btnSchLabel(iBtnIndex).Height = vsfTime.RowHeight(1)
        btnSchLabel(iBtnIndex).BackColor = lngColor
        iPreIndex = iBtnIndex
        '�������2��Сʱ����ѭ���ڷ��м�����ı�ǩ
        For i = 1 To IIf(intEMinute = 0, intEHour - intSHour - 2, intEHour - intSHour - 1)
            '������ױ�ǩԭ���Ѿ��б�ǩ�飬�������㹻�������б�ǩ, ��ֱ��ʹ�����еı�ǩ��û���򴴽��±�ǩ
            If btnSchLabel(iPreIndex).HelpContextID = 0 _
                Or (btnSchLabel(iPreIndex).HelpContextID = iBtnIndex) Then
                iNewIndex = CreateNewSchLabel()
            Else
                iNewIndex = btnSchLabel(iPreIndex).HelpContextID
            End If
            
            '���ñ�ǩ��λ��
            btnSchLabel(iNewIndex).Left = vsfTime.ColWidth(0)
            btnSchLabel(iNewIndex).Top = lngSY + i * vsfTime.RowHeight(1)
            btnSchLabel(iNewIndex).Width = vsfTime.Width - vsfTime.ColWidth(0)
            btnSchLabel(iNewIndex).Height = vsfTime.RowHeight(1)
            '���ñ�ǩ�Ļ�����Ϣ
            btnSchLabel(iNewIndex).Caption = btnSchLabel(iBtnIndex).Caption
            btnSchLabel(iNewIndex).tag = btnSchLabel(iBtnIndex).tag
            btnSchLabel(iNewIndex).BackColor = lngColor
            '���ñ�ǩ��
            btnSchLabel(iPreIndex).HelpContextID = iNewIndex
            iPreIndex = iNewIndex
        Next i
        
        '�ڷ����һ�еı�ǩ
        '������ױ�ǩԭ���Ѿ��б�ǩ�飬�������㹻�������б�ǩ, ��ֱ��ʹ�����еı�ǩ��û���򴴽��±�ǩ
        If btnSchLabel(iPreIndex).HelpContextID = 0 _
            Or (btnSchLabel(iPreIndex).HelpContextID = iBtnIndex) Then
            iNewIndex = CreateNewSchLabel()
        Else
            iNewIndex = btnSchLabel(iPreIndex).HelpContextID
            '����ԭ����ǩ������Ķ����ǩ������Ϊ���ɼ�����ж��
            If btnSchLabel(iNewIndex).HelpContextID <> iBtnIndex Then
                Call DelUnUseLabel(iBtnIndex, btnSchLabel(iNewIndex).HelpContextID)
            End If
        End If
        '���ñ�ǩ��λ��
        btnSchLabel(iNewIndex).Left = vsfTime.ColWidth(0)
        If intEMinute = 0 Then
            btnSchLabel(iNewIndex).Top = lngEY - vsfTime.RowHeight(1)
            btnSchLabel(iNewIndex).Width = vsfTime.Width - vsfTime.ColWidth(0)
        Else
            btnSchLabel(iNewIndex).Top = lngEY
            btnSchLabel(iNewIndex).Width = lngEX - vsfTime.ColWidth(0)
        End If
        
        btnSchLabel(iNewIndex).Height = vsfTime.RowHeight(1)
        '���ñ�ǩ�Ļ�����Ϣ
        btnSchLabel(iNewIndex).Caption = btnSchLabel(iBtnIndex).Caption
        btnSchLabel(iNewIndex).tag = btnSchLabel(iBtnIndex).tag
        btnSchLabel(iNewIndex).BackColor = lngColor
        '���õ�һ����ǩ��HelpContextID���ñ�ǩ���γɱջ�
        btnSchLabel(iPreIndex).HelpContextID = iNewIndex
        btnSchLabel(iNewIndex).HelpContextID = iBtnIndex
    End If
    
    '��ԤԼ��ǩ��������������λ�ã���¼��ԤԼ����
    mSchLabelPool(iPoolIndex).lngBtnCount = IIf(intEMinute = 0, intEHour - intSHour, intEHour - intSHour + 1)
    mSchLabelPool(iPoolIndex).lngRow = GetRowsFromY(btnSchLabel(iBtnIndex).Top)
    mSchLabelPool(iPoolIndex).lngCol = GetColsFromX(btnSchLabel(iBtnIndex).Left)
    
    Call getRowColFromTime(strStartTime, True, lngRow, lngCol)
    lngTimeProjectID = vsfTime.Cell(flexcpData, lngRow, lngCol)
    mSchLabelPool(iPoolIndex).lngTimeProjectID = lngTimeProjectID
    
    'Ĭ�����ÿ�ʼ�ͽ���ʱ��Σ�Ϊ��ʼ�ͽ���ʱ�䣬������¼��ƻ���ԤԼ���ͻ�û�п�ʼ�ͽ���ʱ���
    mSchLabelPool(iPoolIndex).dt��ʼʱ��� = mSchLabelPool(iPoolIndex).dtStartTime
    mSchLabelPool(iPoolIndex).dt����ʱ��� = mSchLabelPool(iPoolIndex).dtEndTime
    If lngTimeProjectID <> 0 Then
        For i = 1 To UBound(mSchTimeProject)
            If mSchTimeProject(i).lngID = lngTimeProjectID Then
                mSchLabelPool(iPoolIndex).dt��ʼʱ��� = Format(mSchLabelPool(iPoolIndex).dtStartTime, "YYYY-MM-DD") & " " & Format(mSchTimeProject(i).dtStartTime, "HH:MM:SS")
                mSchLabelPool(iPoolIndex).dt����ʱ��� = Format(mSchLabelPool(iPoolIndex).dtStartTime, "YYYY-MM-DD") & " " & Format(mSchTimeProject(i).dtEndTime, "HH:MM:SS")
            End If
        Next i
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function GetSchPlanID(ByVal lngSchDeviceID As Long, ByRef dtDate As Date, _
    ByVal blnFindNextDay As Boolean, ByVal blnSilent As Boolean) As Long
'------------------------------------------------
'���ܣ������豸ID�����ڣ����Ҵ����쿪ʼ����ӽ�һ��Ŀ��÷���ID
'������ lngSchDeviceID -- ԤԼ�豸ID
'       dtDate -- �����ز�����ԤԼ���ڡ��Ȳ���dtDate����ķ���ID�����blnFindNextDay=True����û�п��÷������Զ�������һ����ԤԼ����������
'       blnFindNextDay -- ���dtDateû�п��õ�ԤԼ�������Ƿ������һ�����õ�ԤԼ����ID
'       blnSilent -- �Ƿ�Ĭ������ʾ�Ի���
'���أ� ԤԼ����ID
'------------------------------------------------
    Dim lngSchPlanID As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim str�������� As String
    Dim strFilter As String
    Dim i As Integer            '��������ֻ����һ��֮�ڵķ���
    
    On Error GoTo err
    
    '����ԤԼ�豸ID�����ڣ����Ҷ�Ӧ��ʱ�䷽��ID
    strSQL = "select ID,��������,��������,��������,�Ƿ�����,��ʼʱ��,���,�Ƿ�������Ϣ " _
        & " from Ӱ��ԤԼ���� where ԤԼ�豸ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�����ԤԼ����", lngSchDeviceID)
    
    i = 0
    If rsTemp.EOF = False Then
        '�Ȳ����Ƿ���
        Do
            '�Ȳ��ҽ����һ�췽��
            strFilter = "��������='" & Format(dtDate, "YYYYMMDD") & "'"
            rsTemp.Filter = strFilter
            
            If rsTemp.EOF = False Then
                lngSchPlanID = rsTemp!ID
                blnFindNextDay = False  '�ҵ�һ��һ�췽�����˳�ѭ��
            Else
                '�ٲ����Ѿ����õķ���
                strFilter = "�Ƿ�����=1 "
                rsTemp.Filter = strFilter
                If rsTemp.EOF = False Then
                    '�����Ƿ��к��ʽ���ķ��������� ��ÿ�족����ÿ�ܡ�����ÿ�¡���˳���ж�
                    If rsTemp!�������� = Sch_PlanType_ÿ�� And rsTemp!��ʼʱ�� < dtDate Then
                        If rsTemp!��� <> 0 Then
                            If DateDiff("d", rsTemp!��ʼʱ��, dtDate) Mod rsTemp!��� = 0 Then
                                lngSchPlanID = rsTemp!ID
                                blnFindNextDay = False  '�ҵ�һ�����÷������˳�����
                            End If
                        Else
                            lngSchPlanID = rsTemp!ID
                            blnFindNextDay = False  '�ҵ�һ�����÷������˳�����
                        End If
                    ElseIf rsTemp!�������� = Sch_PlanType_ÿ�� And rsTemp!��ʼʱ�� < dtDate Then
                        '��ÿ�ܡ�������Ҫ�жϽ������ܼ�
                        '����ж��ÿ�ܷ�������ҪѰ���ʺ�dtDate��ÿ�ܷ���
                        '���ʹ��ԤԼ��������Ϣ�ղ���ԤԼ
                        rsTemp.MoveFirst
                        While rsTemp.EOF = False
                            str�������� = rsTemp!��������
                            If nvl(rsTemp!�Ƿ�������Ϣ, 0) = 1 And IsDayOff(dtDate) = True Then
                                'ʲô������
                            Else
                                If InStr(str��������, Weekday(dtDate, vbMonday)) > 0 Then
                                    If rsTemp!��� <> 0 Then
                                        If DateDiff("w", rsTemp!��ʼʱ��, dtDate) Mod rsTemp!��� = 0 Then
                                            lngSchPlanID = rsTemp!ID
                                            blnFindNextDay = False  '�ҵ�һ�����÷������˳�����
                                        End If
                                    Else
                                        lngSchPlanID = rsTemp!ID
                                        blnFindNextDay = False  '�ҵ�һ�����÷������˳�����
                                    End If
                                End If
                            End If
                            rsTemp.MoveNext
                        Wend
                    Else    'ÿ�·�������ѯԤԼ����
                        If IsDayOff(dtDate) = False Then
                            lngSchPlanID = rsTemp!ID
                            blnFindNextDay = False  '�ҵ�һ�����÷������˳�����
                        End If
                    End If
                    
                ElseIf blnFindNextDay = False Then
                    lngSchPlanID = 0
                End If
            End If
            
            'û���ҵ�ԤԼ������������һ��
            If blnFindNextDay = True Then
                dtDate = dtDate + 1
                i = i + 1
            End If
        Loop Until (blnFindNextDay = False) Or i > 365
    Else
        lngSchPlanID = 0
    End If
    
    If lngSchPlanID = 0 And blnSilent = False Then
        MsgBox dtDate & " û�п��õ�ԤԼ��������������ԤԼ������", vbOKOnly, "���ԤԼ��ʾ"
    End If
    GetSchPlanID = lngSchPlanID
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function FindFirstSchDay(ByVal lngSchDeviceID As Long, ByRef dtSchDate As Date) As Long
'------------------------------------------------
'���ܣ������豸ID�����ڣ����Ҵ����쿪ʼ���󣬵�һ����ԤԼ����������ԤԼ������
'������ lngSchDeviceID -- ԤԼ�豸ID
'       dtSchDate -- ���ز�������ӽ� dtSchDate ��ԤԼ���ڡ�
'���أ� ԤԼ����ID
'------------------------------------------------
    Dim lngSchPlanID As Long        'ԤԼ����ID
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim lngSchCapacity As Long      'ԤԼ����
    Dim blnFindNext As Boolean      '������һ������
    Dim i As Integer                '��������ֻ��ѯһ��֮�ڵĿ�ԤԼ����
    
    On Error GoTo err
    
    blnFindNext = True
    FindFirstSchDay = 0
    i = 0
    
    While blnFindNext = True And i < 365
        '���ȸ���ԤԼ�豸ID �� ʱ�䣬���ҵ�һ������ԤԼ�����ں�ԤԼ����ID
        lngSchPlanID = GetSchPlanID(lngSchDeviceID, dtSchDate, True, False)
        If lngSchPlanID = 0 Then
            Exit Function
        End If
        
        '�鿴ԤԼ�����Ƿ���ԤԼ����������У�ֱ��ԤԼ�����û�У������ҵ������һ��ԤԼ
        strSQL = "select sum(ԤԼ����) as ����,count(id) as ����  from Ӱ��ԤԼʱ��ƻ� where ԤԼ����ID = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯԤԼ����", lngSchPlanID)
        
        If nvl(rsTemp!����, 0) = 0 = True Then
            MsgBox "�޷�ԤԼ����ԤԼ����û��ʱ��ƻ����������ú�ԤԼʱ��ƻ���ԤԼ����ID=" & lngSchPlanID, vbOKOnly, "���ԤԼ��ʾ"
            Exit Function
        End If
        
        If nvl(rsTemp!����, 0) = 0 Then
            MsgBox "�޷�ԤԼ��ʱ��ƻ��У�ԤԼ����Ϊ0������ϵ����Ա��������ԤԼ������", vbOKOnly, "���ԤԼ��ʾ"
            Exit Function
        End If
        
        lngSchCapacity = rsTemp!����
        
        strSQL = "select " & lngSchCapacity & "- count(ID) as ʣ������ from Ӱ��ԤԼ��¼ where ԤԼ�豸ID=[1] and ԤԼ��ʼʱ�� between [2] and [3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���ԤԼ����", mlngSchDeviceID, CDate(Format(dtSchDate, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(dtSchDate, "yyyy-mm-dd") & " 23:59:59"))
        
        If rsTemp!ʣ������ > 0 Then
            '�������ԤԼ
            '����ǽ��죬�Ǿͻ���Ҫ�����жϣ���ǰʱ��֮���Ƿ���ԤԼ�ƻ�
            If Format(dtSchDate, "YYYY-MM-DD") = Format(Now, "YYYY-MM-DD") Then
                strSQL = "Select Sum(a.ԤԼ����) As ���� From Ӱ��ԤԼʱ��ƻ� A " _
                    & " Where a.ԤԼ����ID = [1] And to_char(a.����ʱ��, 'hh24:mi:ss') > to_char(sysdate+ 2 / 24, 'hh24:mi:ss')"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�жϽ����ʱ�Ƿ���ԤԼ����", lngSchPlanID)
                If nvl(rsTemp!����, 0) = 0 Then
                    '������һ������
                    dtSchDate = dtSchDate + 1
                    i = i + 1
                Else
                    FindFirstSchDay = lngSchPlanID
                    blnFindNext = False     '�˳�ѭ��
                End If
            Else
                FindFirstSchDay = lngSchPlanID
                blnFindNext = False     '�˳�ѭ��
            End If
        Else
            '������һ������
            dtSchDate = dtSchDate + 1
            i = i + 1
        End If
    Wend
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CreateNewSchLabel() As Long
'------------------------------------------------
'���ܣ�����һ���µ�ԤԼ��ǩ
'��������
'���أ���ԤԼ��ǩ������
'------------------------------------------------
    On Error GoTo err
    
    mlngSchLabelIndex = mlngSchLabelIndex + 1
    Load btnSchLabel(mlngSchLabelIndex)
    btnSchLabel(mlngSchLabelIndex).Visible = True
    btnSchLabel(mlngSchLabelIndex).ZOrder
    CreateNewSchLabel = mlngSchLabelIndex
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetRowsFromY(ByVal lngY As Long) As Long
'------------------------------------------------
'���ܣ����ݵ�ǰ��Yλ�ã��������ʱ����е�����
'������ lngY --- ��ǰ������ڵ�Yλ��
'���أ���ʱ����ڣ���ӽ����λ�õ�����
'------------------------------------------------
    Dim i As Integer
    
    On Error GoTo err
        
    '����������У����Զ��ŵ���һ��
    If lngY <= vsfTime.RowHeight(0) Then
        GetRowsFromY = 1
    ElseIf lngY >= vsfTime.Height Then
        GetRowsFromY = 24
    Else
        For i = 0 To 23
            If lngY < vsfTime.RowHeight(0) + vsfTime.RowHeight(1) * i Then
                Exit For
            End If
        Next i
        GetRowsFromY = i
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetColsFromX(ByVal lngX As Long) As Long
'------------------------------------------------
'���ܣ����ݵ�ǰ��Xλ�ã��������ʱ����е�����
'������ lngX --- ��ǰ������ڵ�Xλ��
'���أ���ʱ����ڣ���ӽ����λ�õ�����
'------------------------------------------------
    Dim i As Integer
    
    On Error GoTo err
        
    '����������У���Żص�һ��
    If lngX < vsfTime.ColWidth(0) Or lngX >= vsfTime.Width Then
        GetColsFromX = 0
    Else
        For i = 0 To 10
            If lngX < vsfTime.ColWidth(0) + vsfTime.ColWidth(1) * i Then
                Exit For
            End If
        Next i
        GetColsFromX = i
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ResortSchLabelPool(iPoolIndex As Integer) As Long
'------------------------------------------------
'���ܣ����������ӵ���ţ����¸�ԤԼ������ԤԼ���ڣ������������
'������ iPoolIndex --- ��ź�λ�÷����ı�ı�ǩ����ԤԼ���ڵ�����
'���أ�ֱ����ԤԼ���ڣ���iPoolIndex��λ�������򣬷���ֵ���µ����
'------------------------------------------------
    Dim iSwapIndex As Integer
    Dim i As Integer
    Dim tmpSchLabel As TYPE_SchLabel
    
    On Error GoTo err
        
    '������Ĭ��ֵ����ŵ�λ��û�з����ı�
    ResortSchLabelPool = iPoolIndex
    
    For i = 1 To UBound(mSchLabelPool)
        If i <> iPoolIndex Then
            If mSchLabelPool(iPoolIndex).lng��� > mSchLabelPool(i).lng��� Then
                '��
            Else
                Exit For
            End If
        End If
    Next i
    iSwapIndex = i
    If iPoolIndex < iSwapIndex Then iSwapIndex = iSwapIndex - 1
    
    If iSwapIndex <> 0 And iSwapIndex <> iPoolIndex Then
        '�����С������Ҫ�������򣬽�iPoolIndexָ������ݣ�����iSwapIndex��λ��
        tmpSchLabel = mSchLabelPool(iPoolIndex)
        If iPoolIndex < iSwapIndex Then
            '�����ƶ���ǩ����������ǩ��Ҫ��ǰ��
            For i = iPoolIndex To iSwapIndex - 1
                '��ǰ�ƶ�ԤԼ�ؿ��λ��
                mSchLabelPool(i) = mSchLabelPool(i + 1)
                '��������ԤԼ�غ�ԤԼ��ǩ�Ĺ�ϵ
                Call setSchLabelTag(mSchLabelPool(i).lngBtnIndex, i)
            Next i
        Else
            '��ǰ�ƶ���ǩ����������ǩ��Ҫ������
            For i = iPoolIndex To iSwapIndex + 1 Step -1
                '����ƶ�ԤԼ�ؿ��λ��
                mSchLabelPool(i) = mSchLabelPool(i - 1)
                '��������ԤԼ�غ�ԤԼ��ǩ�Ĺ�ϵ
                Call setSchLabelTag(mSchLabelPool(i).lngBtnIndex, i)
            Next i
        End If
        
        mSchLabelPool(iSwapIndex) = tmpSchLabel
        Call setSchLabelTag(mSchLabelPool(iSwapIndex).lngBtnIndex, iSwapIndex)
    End If
    
    ResortSchLabelPool = iSwapIndex
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetTimeFromXY(lngX As Long, lngY As Long) As Date
'------------------------------------------------
'���ܣ�����X,Y��λ�ã�������ʱ����ڶ�Ӧ��ʱ��
'������ lngX --- Xλ��
'       lngY --- Yλ��
'���أ���Ӧ��ʱ��
'------------------------------------------------
    Dim lngMinute As Long
    Dim lngHour As Long
    Dim i As Integer
    
    On Error GoTo err
    
    '����Y���õ����ڵ��У�����Сʱ
    lngHour = GetRowsFromY(lngY) - 1
    
    '����X��ʱ����еõ��������õ�����
    '����������У���Ϊ0
    If lngX <= vsfTime.ColWidth(0) Then
        lngMinute = 0
    ElseIf lngX >= vsfTime.Width Then
        lngMinute = 60  '������ڿ�ȣ���Ҫ��Сʱ�����λ
    Else
        For i = 0 To 10
            If lngX < vsfTime.ColWidth(0) + vsfTime.ColWidth(1) * i Then
                Exit For
            End If
        Next i
        i = i - 1
        lngMinute = i * 5 + ((lngX - (vsfTime.ColWidth(0) + vsfTime.ColWidth(1) * i)) / vsfTime.ColWidth(1) * 5)
    End If
    
    
    '������ڵ���60,��Ҫ��Сʱ�����λ
    If lngMinute >= 60 Then
        lngMinute = 0
        lngHour = lngHour + 1
    End If
    
    GetTimeFromXY = lngHour & ":" & lngMinute
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetXYFromTime(ByVal strTime As String, ByRef lngX As Long, ByRef lngY As Long) As Boolean
'------------------------------------------------
'���ܣ�����ʱ����ڵ�ʱ�䣬�����Ӧ��X,Yλ��
'������ strTime --- ʱ����ڵ�ʱ�䣬��ʽΪ��HH:MM��
'       lngX --- ���ز�����Xλ��
'       lngY --- ���ز�����Yλ��
'���أ���Ӧ��ʱ��
'------------------------------------------------
    Dim lngMinute As Long
    Dim lngHour As Long
    Dim lngRow As Long
    Dim lngCol As Long
    
    On Error GoTo err
    
    If UBound(Split(strTime, ":")) <> 1 Then
        GetXYFromTime = False
        Exit Function
    End If
    
    lngHour = Split(strTime, ":")(0)
    lngMinute = Split(strTime, ":")(1)
    
    lngRow = lngHour + 1
    lngY = vsfTime.RowHeight(0) + vsfTime.RowHeight(1) * (lngRow - 1)
    
    lngCol = Int(lngMinute / 5 + 1) 'ȷ��5�ı�����������һ��
    lngX = vsfTime.ColWidth(0) + vsfTime.ColWidth(1) * (lngCol - 1) + (lngMinute Mod 5) / 5 * vsfTime.ColWidth(1)
    
    GetXYFromTime = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

'��ʱ������ȷ�Ϲ���ȫ��������ɾ��
'Private Function MoveBtnLabelsAuto(iBtnIndex As Integer, lngBaseX As Long, lngBaseY As Long, ByVal lngX As Long, ByVal lngY As Long) As Boolean
''------------------------------------------------
''���ܣ�������ǩ���λ��
''������ iBtnIndex --- �ƶ���ǩ����
''       lngBaseX --- �ƶ�ǰ��Xλ��
''       lngBaseY --- �ƶ�ǰ��Yλ��
''       lngX --- �ƶ����Xλ��
''       lngY --- �ƶ����Yλ��
''���أ�True--�ɹ���False--ʧ��
''------------------------------------------------
'    Dim lngLeft As Long
'    Dim lngTop As Long
'    Dim iFirstIndex As Integer      '��ǩ���е�һ����ǩ������
'    Dim iLastIndex As Integer       '��ǩ�������һ����ǩ������
'    Dim dtStartTime As Date         '��ʼʱ��
'    Dim dtEndTime As Date           '����ʱ��
'    Dim lngRight As Long
'    Dim dtStartTimeSQL As Date
'    Dim dtEndTimeSQL As Date
'    Dim blnFind As Boolean '�ҵ�ʱ���
'    Dim blnTimeOK As Boolean 'ʱ���������
'    Dim blnNeedTestTime As Boolean '��Ҫ��֤ʱ��
'    Dim tmTmp As Date
'
'    Dim intPoolIndex As Integer
'
'
'    Dim i As Integer, j As Integer
'
'    On Error GoTo err
'
'    blnFind = False
'    blnNeedTestTime = False
'    '�ƶ���ǩ�ļ��������
'    '1���ƶ�������ǩ��������
'    '2���ƶ�������ǩ����ʼ���䣬���ϻ������¹���
'    '3���ƶ���ǩ��
'
'    '�����ǩ�Ĵ������ԤԼ��ǩ���������������ҿ�ȣ����Զ�����һ���µ�ԤԼ��ǩ���γ�һ�ױ�ǩ��
'    '��ǩ��ʹ�� HelpContexID��Ϊ���ӱ�ǣ����ڵı�ǩ�������¼HelpContextID���γɱջ�
'    '��ǩ���е�һ����ǩ����������¼�� ��ǩ�� mSchLabelPool �� lngBtnIndex ��
'    '��ǩ�����������¼�ڱ�ǩ�� mSchLabelPool �� lngBtnCount ��
'    '�����ж��Ƿ��漰����ǩ���������
'
'    If btnSchLabel(iBtnIndex).HelpContextID <> 0 Then   '�Ǳ�ǩ��
'        '�����ҵ������ǩ���еĵ�һ����ǩ����
'        iFirstIndex = mSchLabelPool(btnSchLabel(iBtnIndex).tag).lngBtnIndex
'        iLastIndex = iFirstIndex
'        While btnSchLabel(iLastIndex).HelpContextID <> iFirstIndex
'            iLastIndex = btnSchLabel(iLastIndex).HelpContextID
'        Wend
'
'    Else    '���Ǳ�ǩ�飬��ǰ�������ǵ�һ������
'        iFirstIndex = iBtnIndex
'        iLastIndex = iBtnIndex
'    End If
'
'    intPoolIndex = btnSchLabel(iBtnIndex).tag
'
'    '�����һ����ǩ����λ��
'    lngLeft = btnSchLabel(iFirstIndex).Left + (lngX - lngBaseX)
'    lngTop = btnSchLabel(iFirstIndex).Top + (lngY - lngBaseY)
'
'    '�жϴ˱�ǩ�Ƿ�ᳬ��ʱ���Χ
'    '���������ƶ��Ĺ����У���ǩ������ʱ���Χ����ȡ�������ƶ����ñ�ǩ�����ڵ�ǰλ��
'    If lngTop < vsfTime.RowHeight(0) Or lngTop > vsfTime.Height Then
'        '��ǩ�Ѿ����Ϸ����·�������ʱ���ֹͣ�ƶ���ǩ
'        Exit Function
'    ElseIf lngLeft < vsfTime.ColWidth(0) Then
'        lngLeft = vsfTime.ColWidth(0)
'    ElseIf lngLeft > vsfTime.Width Then
'        '��ǩ���ұ߳���ʱ�����Ҫ�����ƶ�һ��
'        lngLeft = vsfTime.ColWidth(0) + (lngLeft - vsfTime.Width)
'        lngTop = lngTop + vsfTime.RowHeight(1)
'    End If
'
'    '�����µĿ�ʼʱ��
'    dtStartTime = GetTimeFromXY(lngLeft, lngTop)
'    dtEndTime = DateAdd("n", DateDiff("n", mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtStartTime, mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtEndTime), dtStartTime)
'    If DateDiff("d", dtStartTime, dtEndTime) = 1 Then
'        '�±�ǩ����ʱ������ޣ���ֹ�ƶ���ֱ���˳�
'        Exit Function
'    End If
'
'    '���ݿ�ʼʱ���жϴ�����һ��ʱ��Σ����Ҹ���ԤԼ�ص����µĿ�ʼ
'
''   '�������λ���жϵ�ǰ�����ĸ�ʱ���
'    For i = 1 To UBound(mSchTimeProject)
'         If mSchTimeProject(i).dtStartTime <= dtStartTime And mSchTimeProject(i).dtEndTime > dtStartTime Then
'            dtStartTimeSQL = mSchTimeProject(i).dtStartTime
'            dtEndTimeSQL = mSchTimeProject(i).dtEndTime
'            blnFind = True
'            Exit For
'         End If
'    Next
'
'    If Not blnFind Then
'        dtStartTimeSQL = mSchTimeProject(1).dtStartTime
'        dtEndTimeSQL = mSchTimeProject(1).dtEndTime
'    End If
'
'    blnNeedTestTime = True
'
'    blnTimeOK = True
'    '��������ǵ�ǰ��ѡ�е�ʱ���ǡ�ÿ�Խ��ǰʱ�䣬���Ե�ǰʱ����Ϊ��ʼʱ���ж��Ƿ����������������ϣ�ֱ������Ϊ�Զ����õ�ʱ��
'    If Format(Now, "YYYY-MM-DD") = Format(mdtSchDate, "YYYY-MM-DD") Then
'        tmTmp = Format(Now, "hh:mm:ss")
'        tmTmp = DateAdd("n", 5, tmTmp)
'        If dtStartTimeSQL <= tmTmp And dtEndTimeSQL >= tmTmp Then
'            dtStartTime = tmTmp
'
'            dtEndTime = DateAdd("n", DateDiff("n", mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtStartTime, mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtEndTime), dtStartTime)
'            For i = 1 To UBound(mSchLabelPool) - 1
'                '�ܿ���ǰʱ���
'                If intPoolIndex <> i Then '��Ҫ�޸��������
'                    If Not ((Format(mSchLabelPool(i).dtStartTime, "hh:mm:ss") < dtStartTime And Format(mSchLabelPool(i).dtEndTime, "hh:mm:ss") <= dtStartTime) Or (Format(mSchLabelPool(i).dtStartTime, "hh:mm:ss") >= dtEndTime And Format(mSchLabelPool(i).dtEndTime, "hh:mm:ss") > dtEndTime)) Then
'                        blnTimeOK = False
'                    End If
'                End If
'            Next
'        End If
'    End If
'
'    If UBound(mSchLabelPool) = 1 Then
'        If Not blnTimeOK Then
'            dtStartTime = dtStartTimeSQL
'            dtEndTime = dtEndTimeSQL
'        End If
'    Else
'
'        If mSchLabelPool(1).dtStartTime >= Format(mdtSchDate, "YYYY-MM-DD") & dtEndTimeSQL Then
'            '���������˳�whileѭ��
'            dtStartTime = dtStartTimeSQL
'            dtEndTime = dtEndTimeSQL
'            blnNeedTestTime = False
'        End If
'
'        '������Ȼ����ʱ��ο�ʼʱ����ΪԤԼ��ʼʱ����Ч
'
'        dtStartTime = dtStartTimeSQL
'        dtEndTime = DateAdd("n", DateDiff("n", mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtStartTime, mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtEndTime), dtStartTime)
'
'
'        If Not blnTimeOK Then
'            For i = 1 To UBound(mSchLabelPool) - 1
'                '�ܿ���ǰʱ���
'                If intPoolIndex <> i Then '��Ҫ�޸��������
'                    If Not ((Format(mSchLabelPool(i).dtStartTime, "hh:mm:ss") < dtStartTime And Format(mSchLabelPool(i).dtEndTime, "hh:mm:ss") <= dtStartTime) Or (Format(mSchLabelPool(i).dtStartTime, "hh:mm:ss") >= dtEndTime And Format(mSchLabelPool(i).dtEndTime, "hh:mm:ss") > dtEndTime)) Then
'                        blnTimeOK = False
'                    End If
'                End If
'            Next
'        End If
'
'        If Not blnTimeOK Then
'
'            For i = 1 To UBound(mSchLabelPool) - 1
'                If blnNeedTestTime And Format(mSchLabelPool(i).dtStartTime, "hh:mm:ss") >= dtStartTimeSQL And Format(mSchLabelPool(i).dtStartTime, "hh:mm:ss") <= dtEndTimeSQL And iBtnIndex <> i Then
'                    blnTimeOK = True
'
'                    If blnTimeOK And blnNeedTestTime Then
'                        dtStartTime = Format(mSchLabelPool(i).dtEndTime, "hh:mm:ss") 'rsSchedule!ԤԼ����ʱ�� 'Format(rsSchedule!ԤԼ����ʱ��, "hh-mm-ss")
'                        dtEndTime = DateAdd("n", DateDiff("n", mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtStartTime, mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtEndTime), dtStartTime)
'
'                        For j = 1 To UBound(mSchLabelPool)
'                            If intPoolIndex <> j Then '��Ҫ�޸��������
'                                If Format(mSchLabelPool(i).dtStartTime, "hh:mm:ss") >= dtStartTimeSQL And Format(mSchLabelPool(i).dtStartTime, "hh:mm:ss") <= dtEndTimeSQL Then
'                                    If Not ((Format(mSchLabelPool(j).dtStartTime, "hh:mm:ss") < dtStartTime And Format(mSchLabelPool(j).dtEndTime, "hh:mm:ss") <= dtStartTime) Or (Format(mSchLabelPool(j).dtStartTime, "hh:mm:ss") >= dtEndTime And Format(mSchLabelPool(j).dtEndTime, "hh:mm:ss") > dtEndTime)) Then
'                                        blnTimeOK = False
'                                    End If
'                                End If
'                            End If
'                        Next
'
'                        If blnTimeOK Then
'                            blnNeedTestTime = False
'                        End If
'                    End If
'                End If
'            Next
'        End If
'    End If
'
'    Call ModifySchPoolTime(btnSchLabel(iFirstIndex).tag, dtStartTime)
'    '���°ڷű�ǩ
'    Call PutSchLabel(iFirstIndex, btnSchLabel(iFirstIndex).tag)
'
'    If iFirstIndex <> iBtnIndex Then
'        mlngBaseX = lngLeft
'    End If
'
'    Exit Function
'err:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'End Function

Private Function MoveBtnLabels(iBtnIndex As Integer, lngBaseX As Long, lngBaseY As Long, ByVal lngX As Long, ByVal lngY As Long) As Boolean
'------------------------------------------------
'���ܣ�������ǩ���λ��
'������ iBtnIndex --- �ƶ���ǩ����
'       lngBaseX --- �ƶ�ǰ��Xλ��
'       lngBaseY --- �ƶ�ǰ��Yλ��
'       lngX --- �ƶ����Xλ��
'       lngY --- �ƶ����Yλ��
'���أ�True--�ɹ���False--ʧ��
'------------------------------------------------
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim iFirstIndex As Integer      '��ǩ���е�һ����ǩ������
    Dim iLastIndex As Integer       '��ǩ�������һ����ǩ������
    Dim dtStartTime As Date         '��ʼʱ��
    Dim dtEndTime As Date           '����ʱ��
    Dim lngRight As Long
    
    Dim i As Integer
    
    On Error GoTo err
    
    '�ƶ���ǩ�ļ��������
    '1���ƶ�������ǩ��������
    '2���ƶ�������ǩ����ʼ���䣬���ϻ������¹���
    '3���ƶ���ǩ��
    
    '�����ǩ�Ĵ������ԤԼ��ǩ���������������ҿ�ȣ����Զ�����һ���µ�ԤԼ��ǩ���γ�һ�ױ�ǩ��
    '��ǩ��ʹ�� HelpContexID��Ϊ���ӱ�ǣ����ڵı�ǩ�������¼HelpContextID���γɱջ�
    '��ǩ���е�һ����ǩ����������¼�� ��ǩ�� mSchLabelPool �� lngBtnIndex ��
    '��ǩ�����������¼�ڱ�ǩ�� mSchLabelPool �� lngBtnCount ��
    
    '�����ж��Ƿ��漰����ǩ���������
    If btnSchLabel(iBtnIndex).HelpContextID <> 0 Then   '�Ǳ�ǩ��
        '�����ҵ������ǩ���еĵ�һ����ǩ����
        iFirstIndex = mSchLabelPool(btnSchLabel(iBtnIndex).tag).lngBtnIndex
        iLastIndex = iFirstIndex
        While btnSchLabel(iLastIndex).HelpContextID <> iFirstIndex
            iLastIndex = btnSchLabel(iLastIndex).HelpContextID
        Wend
        
    Else    '���Ǳ�ǩ�飬��ǰ�������ǵ�һ������
        iFirstIndex = iBtnIndex
        iLastIndex = iBtnIndex
    End If
        
    '�����һ����ǩ����λ��
    lngLeft = btnSchLabel(iFirstIndex).Left + (lngX - lngBaseX)
    lngTop = btnSchLabel(iFirstIndex).Top + (lngY - lngBaseY)

    '�жϴ˱�ǩ�Ƿ�ᳬ��ʱ���Χ
    '���������ƶ��Ĺ����У���ǩ������ʱ���Χ����ȡ�������ƶ����ñ�ǩ�����ڵ�ǰλ��
    If lngTop < vsfTime.RowHeight(0) Or lngTop > vsfTime.Height Then
        '��ǩ�Ѿ����Ϸ����·�������ʱ���ֹͣ�ƶ���ǩ
        Exit Function
    ElseIf lngLeft < vsfTime.ColWidth(0) Then
        '��ǩ����߳���ʱ���ֹͣ�ƶ�
        lngLeft = vsfTime.ColWidth(0) + (lngLeft - vsfTime.Width)
    ElseIf lngLeft > vsfTime.Width Then
        '��ǩ���ұ߳���ʱ�����Ҫ�����ƶ�һ��
        lngLeft = vsfTime.ColWidth(0) + (lngLeft - vsfTime.Width)
'        lngTop = lngTop + vsfTime.RowHeight(1)
    End If
    
    '�����µĿ�ʼʱ��
    dtStartTime = GetTimeFromXY(lngLeft, lngTop)
    dtEndTime = DateAdd("n", DateDiff("n", mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtStartTime, mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtEndTime), dtStartTime)
    If DateDiff("d", dtStartTime, dtEndTime) = 1 Then
        '�±�ǩ����ʱ������ޣ���ֹ�ƶ���ֱ���˳�
        Exit Function
    End If
    '����������е�ʱ��
    Call ModifySchPoolTime(btnSchLabel(iFirstIndex).tag, dtStartTime)
    '���°ڷű�ǩ
    Call PutSchLabel(iFirstIndex, btnSchLabel(iFirstIndex).tag)
    
    '���ֱ���ƶ���ǩ���еڶ����Ժ�ı�ǩ������ʵ������Щ��ǩ��û�б��ƶ��ģ�������Ҫ���ƶ�����һ������
    '���¼�¼mlngBaseX�Ϳ��������ˡ�mlngBaseY����Ҫ��������Ϊ�����Y�����ʵ��λ������0
    If iFirstIndex <> iBtnIndex Then
        mlngBaseX = lngX
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function DelUnUseLabel(ByVal lngBtnIndex As Long, ByVal lngDelIndex As Long) As Boolean
'------------------------------------------------
'���ܣ�ɾ��һ����ǩ���ڣ�����ı�ǩ������ǩ��ı�ǩ��������ʱ����
'������ lngBtnIndex --- ��ǩ���У���һ����ǩ������
'       lngDelIndex --- Ҫ��ʼɾ���ı�ǩ�����������ǩ�Լ��������б�ǩ����ɾ��
'���أ�True -- �ɹ��� False -- ʧ��
'------------------------------------------------
    Dim i  As Integer
    Dim iNextIndex As Integer
    
    On Error GoTo err
    
    
    i = lngDelIndex
    
    Do
        iNextIndex = btnSchLabel(i).HelpContextID

        btnSchLabel(i).Visible = False
        Unload btnSchLabel(i)   '���õı�ǩ��ֱ��ж�ص�
        
        i = iNextIndex
    Loop While (btnSchLabel(i).HelpContextID <> lngBtnIndex) And (btnSchLabel(i).HelpContextID <> 0) And (i <> lngBtnIndex)
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ModifySchPoolTime(ByVal lngPoolIndex As Long, ByVal dtStartTime As Date) As Boolean
'------------------------------------------------
'���ܣ����ݴ���Ŀ�ʼʱ�䣬���µ���ԤԼ���ж�Ӧ�����Ŀ�ʼ�ͽ���ʱ�䣬����ԤԼ����ʱ������
'������ lngPoolIndex --- ԤԼ���е�����
'       dtStartTime --- �µĿ�ʼʱ��
'���أ�True -- �ɹ��� False -- ʧ��
'------------------------------------------------
    Dim lngMinuteDiff As Long   '��ʱ��
    
    On Error GoTo err
    lngMinuteDiff = DateDiff("n", mSchLabelPool(lngPoolIndex).dtStartTime, mSchLabelPool(lngPoolIndex).dtEndTime)
    mSchLabelPool(lngPoolIndex).dtStartTime = Format(mSchLabelPool(lngPoolIndex).dtStartTime, "YYYY-MM-DD") & " " & Format(dtStartTime, "HH:MM:SS")
    mSchLabelPool(lngPoolIndex).dtEndTime = DateAdd("n", lngMinuteDiff, mSchLabelPool(lngPoolIndex).dtStartTime)
    mSchLabelPool(lngPoolIndex).isModified = True
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function AdjustLabelPos(Index As Integer) As Boolean
'------------------------------------------------
'���ܣ���ק�����������̧������ʱ������ԤԼ��ǩ��λ��
'������ Index -- ��ǩ������
'���أ�True -- �ɹ��� False -- ʧ��
'------------------------------------------------
    Dim lngRow As Long
    Dim lngCol As Long
    Dim i As Integer
    Dim iPrePoolIndex As Integer
    Dim iNextPoolIndex As Integer
    Dim lngNewNumber As Long
    Dim iPoolIndex As Integer
    Dim blnIsOutOfTime As Boolean   'ԤԼ��ǩ�Ƿ���ԤԼʱ��ƻ���
    Dim iFirstIndex As Integer
    
    On Error GoTo err
    
    blnIsOutOfTime = False
    
    '�����ж��Ƿ��漰����ǩ���������
    If btnSchLabel(Index).HelpContextID <> 0 Then   '�Ǳ�ǩ��
        '�����ҵ������ǩ���еĵ�һ����ǩ����
        iFirstIndex = mSchLabelPool(btnSchLabel(Index).tag).lngBtnIndex
    Else    '���Ǳ�ǩ�飬��ǰ�������ǵ�һ������
        iFirstIndex = Index
    End If
    iPoolIndex = btnSchLabel(iFirstIndex).tag
    
    '�����ǩ������������ǩ���棬������ʾ����ֹԤԼ
    If IsLabelOverlap(iFirstIndex) = True Then
        '������ʾ��ֱ�ӽ���ǩ�ƶ���ԭ����λ��
        
        mblnRestorePos = True
    End If
    
    '�����ǩ����ק���˷�ԤԼʱ��Σ�������ʾ����������ԤԼ
    If mblnRestorePos = False Then
        lngRow = GetRowsFromY(btnSchLabel(iFirstIndex).Top)
        lngCol = GetColsFromX(btnSchLabel(iFirstIndex).Left)
        
        If (Format(mdtSchDate, "YYYY-MM-DD") = Format(Now, "YYYY-MM-DD")) _
                And (Format(mSchLabelPool(iPoolIndex).dtStartTime, "HH:MM") < Format(Now, "HH:MM")) Then
            MsgBox "�����Ѿ��޷�����飬�����һ��ʱ�������ԤԼ��", vbOKOnly, "���ԤԼ��ʾ"
            mblnRestorePos = True   '����ǩ�ڷŻ�ԭ����λ��
        ElseIf vsfTime.Cell(flexcpData, lngRow, lngCol) = 0 Then
            blnIsOutOfTime = True
            MsgBox "���ڿ���ԤԼ��ʱ����ڣ����ܼ���ԤԼ��", vbOKOnly, "���ԤԼ��ʾ"
            mblnRestorePos = True   '����ǩ�ڷŻ�ԭ����λ��

        End If
    End If
    
    '������ţ���ʼԤԼ
    If mblnRestorePos = False Then
        '����ǩ�������
        lngNewNumber = GetNewNumber(iFirstIndex)
        
        If lngNewNumber = 0 Then
            MsgBox "���ʱ����Ѿ�û�пյ�ԤԼ��ţ��޷�ԤԼ�������һ��ʱ�������ԤԼ��", vbOKOnly, "���ԤԼ��ʾ"
            mblnRestorePos = True   '����ǩ�ڷŻ�ԭ����λ��
        ElseIf lngNewNumber = -1 Then
            mblnRestorePos = True   '����ǩ�ڷŻ�ԭ����λ��
        End If
        
        '���������ԤԼ
        If mblnRestorePos = False Then
            mSchLabelPool(iPoolIndex).lng��� = lngNewNumber
            '���µ���ԤԼ��˳��
            Call ResortSchLabelPool(iPoolIndex)
        End If
    End If
    
     '�����ǩ����ק����ʱ���֮�⣬��Ҫ�ָ���ǩ��ԭ��λ��
    If mblnRestorePos = True Then
        Call RestoreLabelPos(iFirstIndex)
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function IsLabelOverlap(Index As Integer) As Boolean
'------------------------------------------------
'���ܣ��жϵ�ǰ��ǩ���Ƿ��������ǩ��ʱ���ϸ�����
'������ Index --- ��ǩ������
'���أ�True -- �и��ǣ� False -- �޸���
'------------------------------------------------
    Dim i As Integer
    Dim strS1 As String
    Dim strS2 As String
    Dim strE1 As String
    Dim strE2 As String
    Dim iFirstIndex As Integer
    
    On Error GoTo err
    
    iFirstIndex = mSchLabelPool(btnSchLabel(Index).tag).lngBtnIndex
    strS2 = Format(mSchLabelPool(btnSchLabel(Index).tag).dtStartTime, "HH:MM")
    strE2 = Format(mSchLabelPool(btnSchLabel(Index).tag).dtEndTime, "HH:MM")
    
    '����ʱ�����ж��Ƿ���ڱ�ǩ����
    For i = 1 To UBound(mSchLabelPool)
        If mSchLabelPool(i).lngBtnIndex <> iFirstIndex Then
            strS1 = Format(mSchLabelPool(i).dtStartTime, "HH:MM")
            strE1 = Format(mSchLabelPool(i).dtEndTime, "HH:MM")
            
            If (strS1 <= strS2 And strS2 < strE1) _
                Or (strS1 < strE2 And strE2 <= strE1) _
                Or (strS2 <= strS1 And strS1 < strE2) Then
                IsLabelOverlap = True
                Exit Function
            End If
        End If
    Next i
    
    IsLabelOverlap = False
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetNewNumber(iBtnIndex As Integer) As Long
'------------------------------------------------
'���ܣ����ݵ�ǰԤԼ��ǩ���ڵ�λ�ã�������µ�ԤԼ���
'������ iBtnIndex --- ԤԼ��ǩ������
'���أ��µ�ԤԼ���,-1��ʾ�û�����ֹͣԤԼ��0��ʾû����ţ������������
'------------------------------------------------
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngTimeProjectID As Long
    Dim dtStartTime As Date
    Dim lngBaseNumber As Long       '��ǰʱ��ƻ�֮ǰ��ԤԼ�����ܺͣ���Ż���
    Dim lngNumber As Long           '��������������
    Dim iNumCount As Integer        '���õ��������
    Dim iTimeLength As Integer      'ÿһ�������ռ��Ԥ��ʱ�䳤��
    Dim dtProjectStartTime As Date  'ԤԼʱ��ƻ����У����ʱ��εĿ�ʼʱ��
    Dim dtProjectEndTime As Date    'ԤԼʱ��ƻ����У����ʱ��εĽ���ʱ��
    Dim lngPlanId As Long           'ԤԼʱ��ƻ����У���¼�ķ���ID
    Dim iPoolIndex As Integer       'ԤԼ���е�����
    Dim i As Integer                'ѭ������
    Dim iNextPoolIndex As Integer   '�����֮��ĵ�һ�����
    Dim iPrePoolIndex As Integer    '�����֮ǰ�����һ����ţ����Լ�����Ų�ͬ
    Dim strMsg As String            '��Ϣ
    
    GetNewNumber = 0
    
    On Error GoTo err
    
    iPoolIndex = btnSchLabel(iBtnIndex).tag
    
    '��ԤԼ����ѭ���Աȣ��ҵ���ԤԼ��ŵ�λ�ã�����ʱ����жԱ�
    For i = 1 To UBound(mSchLabelPool)
        If i <> iPoolIndex Then
            If DateDiff("n", mSchLabelPool(iPoolIndex).dtEndTime, mSchLabelPool(i).dtStartTime) >= 0 Then
                Exit For
            End If
        End If
    Next i
    
    '�ҵ���һ������һ��������ȷ����������ڵ�λ��
    iNextPoolIndex = IIf(i > UBound(mSchLabelPool), 0, i)
    If iNextPoolIndex <> 0 And iNextPoolIndex - 1 = iPoolIndex Then
        iPrePoolIndex = iPoolIndex - 1
    ElseIf iNextPoolIndex <> 0 Then
        iPrePoolIndex = iNextPoolIndex - 1
    Else
        iPrePoolIndex = UBound(mSchLabelPool) - 1
    End If
    
    '��������1С�����޺���
    If iNextPoolIndex = 1 And mSchLabelPool(1).lng��� = 1 Then
        Exit Function   'û����ţ��˳�
    ElseIf iNextPoolIndex = 0 And iPoolIndex <> UBound(mSchLabelPool) Then
        '������ƶ������һ����ţ��ҳ��������������ʾ�Ƿ�Ӻţ��Ӻţ������޺���
        '������ƶ������һ����ţ�û�г���ԤԼ��������������
        If mSchLabelPool(UBound(mSchLabelPool)).lng��� >= mlngSchSum Then
            If MsgBox("ԤԼ����Ѿ����������ԤԼ���������Ƿ�����Ӻţ�", vbYesNo, "���ԤԼ��ʾ") = vbYes Then
                GetNewNumber = mSchLabelPool(UBound(mSchLabelPool)).lng��� + 1
                Exit Function   '�ҵ���ţ��˳�
            Else
                GetNewNumber = -1
                Exit Function   'û����ţ��˳�
            End If
        End If
    Else
        '������м���룬ǰ�����֮���޿պţ����޺���
        If mSchLabelPool(iNextPoolIndex).lng��� - mSchLabelPool(iPrePoolIndex).lng��� = 1 Then
            Exit Function   'û����ţ��˳�
        ElseIf mSchLabelPool(iNextPoolIndex).lng��� - mSchLabelPool(iPrePoolIndex).lng��� = 2 Then
            '�պ���һ���պţ���������պ�
            GetNewNumber = mSchLabelPool(iNextPoolIndex).lng��� - 1
            Exit Function   '�ҵ���ţ��˳�
        End If
    End If
    
    '���м����ţ������µ����
    '���ݱ�ǩ��λ�ã������ʱ��ƻ�ID
    lngRow = GetRowsFromY(btnSchLabel(iBtnIndex).Top)
    lngCol = GetColsFromX(btnSchLabel(iBtnIndex).Left)
    lngTimeProjectID = vsfTime.Cell(flexcpData, lngRow, lngCol)
    'lngTimeProjectID = 0 ԤԼ��ǩ����ʱ��ƻ�֮�ڣ�ʹ����һ��ʱ��ε����һ�����룬�����ӾͲ���Ӱ�쵽��һ��ʱ��ε���ű���ˡ�
    
    dtStartTime = mSchLabelPool(iPoolIndex).dtStartTime
    lngBaseNumber = 0
    '���жϵ�ǰ��ǩ����λ���ϣ�ԤԼʱ��ƻ��Ļ�����ʼ���
    For i = 1 To UBound(mSchTimeProject)
        If Format(dtStartTime, "HH:MM:SS") >= Format(mSchTimeProject(i).dtEndTime, "HH:MM:SS") Then
            lngBaseNumber = lngBaseNumber + mSchTimeProject(i).lngSum
        End If
    Next i
    
    '���lngTimeProjectID=0 ,��ʾû����ʱ��ƻ�����ԤԼ����ֱ��ʹ����Ż����е���һ�����
    If lngTimeProjectID = 0 Then
        lngNumber = IIf(lngBaseNumber = 0, 1, lngBaseNumber + 1)  '�����0 ����������1
    Else
        '��ԤԼʱ��ƻ��У���������λ�õ����
        For i = 1 To UBound(mSchTimeProject)
            If mSchTimeProject(i).lngID = lngTimeProjectID Then
                iNumCount = mSchTimeProject(i).lngSum
                dtProjectStartTime = mSchTimeProject(i).dtStartTime
                dtProjectEndTime = mSchTimeProject(i).dtEndTime
                lngPlanId = mSchTimeProject(i).lngSchPlanID
                Exit For
            End If
        Next i
        
        '�ж�ʱ��ƻ��Ƿ��Ѿ�Լ����
        If (SegmentCanUse(dtProjectStartTime, dtProjectEndTime, iNumCount, strMsg)) = 0 Then
            If MsgBox(strMsg & " �Ƿ�����Ӻţ�", vbYesNo, "���ԤԼ��ʾ") = vbNo Then
                GetNewNumber = -1
                Exit Function   'û����ţ��˳�
            End If
        End If
        
        If iNumCount <> 0 Then
            iTimeLength = DateDiff("n", dtProjectStartTime, dtProjectEndTime) / iNumCount
            For i = 1 To iNumCount
                '����ʱ��εĳ��Ⱥ��������������������
                If Format(dtStartTime, "HH:MM") < Format(DateAdd("n", iTimeLength * i, dtProjectStartTime), "HH:MM") Then
                    Exit For
                End If
            Next i
            If i > iNumCount Then
                i = iNumCount
            End If
            lngNumber = lngBaseNumber + i
        End If
    End If
    
    '�������������ţ���Ҫ�ȵ�ǰλ��֮ǰԤԼ��ǩ����Ŵ�
    '�������Ӧ���ǲ����ڵģ��������û��ֹ������˺ܶ��ԤԼ��ǩ�Ŀ�ȣ����ºܶ���žۼ��ڽ϶̵�ʱ����
    If lngNumber <= mSchLabelPool(iPrePoolIndex).lng��� Then
        lngNumber = mSchLabelPool(iPrePoolIndex).lng��� + 1
    ElseIf lngNumber >= mSchLabelPool(iNextPoolIndex).lng��� And iNextPoolIndex <> 0 Then
        lngNumber = mSchLabelPool(iNextPoolIndex).lng��� - 1
    End If
    
    '�ٴ�����Ѱ��һ���յ����
    For i = iPrePoolIndex To UBound(mSchLabelPool)
        If mSchLabelPool(i).lng��� = lngNumber And i <> iPoolIndex Then
            lngNumber = lngNumber + 1
        End If
    Next i
    
    If lngNumber > mlngSchSum Then
        If MsgBox("ԤԼ����Ѿ����������ԤԼ���������Ƿ�����Ӻţ�", vbYesNo, "���ԤԼ��ʾ") = vbNo Then
            GetNewNumber = -1
            Exit Function   'û����ţ��˳�
        End If
    End If
        
    GetNewNumber = lngNumber

    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub SaveAllSchedule()
'------------------------------------------------
'���ܣ��������б��޸Ĺ���ԤԼ��Ϣ
'������
'���أ���
'------------------------------------------------
    Dim i As Integer
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strStartTime As String
    Dim strEndTime As String
    
    On Error GoTo err
    
    For i = 1 To UBound(mSchLabelPool)
        If mSchLabelPool(i).isModified = True Then
            If funSaveSchedule(mSchLabelPool(i).dtStartTime, mSchLabelPool(i).dtEndTime, _
                mSchLabelPool(i).lngҽ��ID, mSchLabelPool(i).str����, mSchLabelPool(i).lng���, _
                mlngSchDeviceID, mSchLabelPool(i).dt��ʼʱ���, mSchLabelPool(i).dt����ʱ���) = False Then
                Exit Sub
            End If
            
            If InStr(mstrModifiedOrderID, CStr(mSchLabelPool(i).lngҽ��ID)) = 0 Then
                mstrModifiedOrderID = mstrModifiedOrderID & "," & CStr(mSchLabelPool(i).lngҽ��ID)
            End If
        End If
    Next i
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SegmentCanUse(dtProjectStartTime As Date, dtProjectEndTime As Date, iCapacity As Integer, ByRef strMsg As String) As Boolean
'------------------------------------------------
'���ܣ��жϵ�ǰʱ����Ƿ����
'������ dtProjectStartTime -- ��ʼʱ��
'       dtProjectEndTime -- ����ʱ��
'       iCapacity -- ԤԼ����
'       strMsg -- ��OUT��ʱ��β�����ʱ������ԭ��
'���أ�True -- ���ã�False -- ������
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim dtStartTime As Date
    Dim dtEndTime As Date
       
    On Error GoTo err
    
    'С�ڽ��죬����False
    If Format(mdtSchDate, "YYYY-MM-DD") < Format(Now, "YYYY-MM-DD") Then
        strMsg = "������ǰ��ʱ�䣬�Ѿ�����ԤԼ��"
        Exit Function
    End If
    
    dtStartTime = CDate(Format(mdtSchDate, "YYYY-MM-DD") & " " & Format(dtProjectStartTime, "hh:mm:ss"))
    dtEndTime = CDate(Format(mdtSchDate, "YYYY-MM-DD") & " " & Format(dtProjectEndTime, "hh:mm:ss"))
    
    '�ǽ��죬�����ǰʱ�䣬�������ʱ�䲻��2Сʱ��Ҳ����False
    If Format(mdtSchDate, "YYYY-MM-DD") = Format(Now, "YYYY-MM-DD") Then
        If DateDiff("n", Now, dtEndTime) <= 120 Then
            strMsg = "���ھ��뱾ʱ��εĽ���ʱ�䲻��2Сʱ��"
            Exit Function
        End If
    End If
    
    '�ж����ʱ��ε�ԤԼ�����Ƿ����ԤԼʱ��ƻ�������
    strSQL = "Select Count(a.���) as SchCount From Ӱ��ԤԼ��¼ A Where a.ԤԼ�豸id = [1] And " _
            & " a.ԤԼ��ʼʱ�� >= [2]  And a.ԤԼ����ʱ�� <=[3] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "ʱ������Ѿ�ԤԼ������", mlngSchDeviceID, _
        CDate(dtStartTime), CDate(dtEndTime))
    If rsTemp!SchCount >= iCapacity Then
        strMsg = "��ǰʱ���ԤԼ����������"
        Exit Function
    End If
        
    SegmentCanUse = True
    
    Exit Function
    
err:
   If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ClearLocalParas()
'------------------------------------------------
'���ܣ����ģ�����
'������
'���أ���
'------------------------------------------------
    
    On Error GoTo err
    
    mlngPoolIndex = 0
    mlngBtnIndex = 0
    mlngOrderID = 0
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub LoadCalendar()
'------------------------------------------------
'���ܣ������ݿ��ȡԤԼ����
'������
'���أ���
'------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo err
    
    strSQL = "select ����,��Ϣ�� from Ӱ��ԤԼ���� where ����>=[1]"
    Set mrsCalendar = zlDatabase.OpenSQLRecord(strSQL, "��ѯԤԼ����", Format(Now, "YYYYMM"))
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function IsDayOff(dtDate As Date) As Boolean
'------------------------------------------------
'���ܣ��ж��Ƿ���Ϣ��
'������ dtDate -- �жϵ�����
'���أ� True -- ����Ϣ�� �� False -- �ǹ�����
'------------------------------------------------
    Dim strFilter As String
    
    On Error GoTo err
    
    IsDayOff = False
    mrsCalendar.Filter = 0
    If mrsCalendar.RecordCount = 0 Then
        Exit Function
    End If
    
    strFilter = "����=" & Format(dtDate, "YYYYMM")
    mrsCalendar.Filter = strFilter
    If mrsCalendar.RecordCount = 1 Then
        If InStr(mrsCalendar!��Ϣ��, Format(dtDate, "DD")) > 0 Then
            IsDayOff = True
        End If
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub setReadOnly()
'------------------------------------------------
'���ܣ������Ƿ�ֻ��ģʽ
'������
'���أ� ��
'------------------------------------------------
        
    On Error GoTo err
    
    vsfTime.Enabled = Not mIsReadOnly
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub setSchLabelTag(iIndex As Long, iPoolIndex As Integer)
'------------------------------------------------
'���ܣ�����ԤԼ��ǩ��TAGֵ������ǹ���Ķ����ǩ��ͬʱ�������б�ǩ��TAGֵ
'������ iIndex -- ԤԼ��ǩ������
'       iPoolIndex -- ԤԼ��ǩ��Ӧ����ص�����
'���أ� ��
'------------------------------------------------
    Dim iTempIndex As Integer
        
    On Error GoTo err
    
    btnSchLabel(iIndex).tag = iPoolIndex
    
    If btnSchLabel(iIndex).HelpContextID <> 0 Then
        iTempIndex = btnSchLabel(iIndex).HelpContextID
        While btnSchLabel(iIndex).HelpContextID <> btnSchLabel(iTempIndex).HelpContextID
            btnSchLabel(iTempIndex).tag = iPoolIndex
            iTempIndex = btnSchLabel(iTempIndex).HelpContextID
        Wend
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub setSchLabelEnable(iIndex As Integer, blnEnable As Boolean)
'------------------------------------------------
'���ܣ�����ԤԼ��ǩ�Ŀ����ԣ�����ǹ���Ķ����ǩ��ͬʱ�������б�ǩ�Ŀ�����
'������ iIndex -- ԤԼ��ǩ������
'       blnEnable -- ��ť�Ƿ����
'���أ� ��
'------------------------------------------------
    Dim iTempIndex As Integer
        
    On Error GoTo err
    
    btnSchLabel(iIndex).Enabled = blnEnable
    
    If btnSchLabel(iIndex).HelpContextID <> 0 Then
        iTempIndex = btnSchLabel(iIndex).HelpContextID
        While btnSchLabel(iIndex).HelpContextID <> btnSchLabel(iTempIndex).HelpContextID
            btnSchLabel(iTempIndex).Enabled = blnEnable
            iTempIndex = btnSchLabel(iTempIndex).HelpContextID
        Wend
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub setSchLabelZorder(iIndex As Integer)
'------------------------------------------------
'���ܣ�����ԤԼ��ǩ��Zorder����ʾ����ǰ��
'������ iIndex -- ԤԼ��ǩ������
'���أ� ��
'------------------------------------------------
    Dim iTempIndex As Integer
        
    On Error GoTo err
    
    Call btnSchLabel(iIndex).ZOrder
    
    If btnSchLabel(iIndex).HelpContextID <> 0 Then
        iTempIndex = btnSchLabel(iIndex).HelpContextID
        While btnSchLabel(iIndex).HelpContextID <> btnSchLabel(iTempIndex).HelpContextID
            Call btnSchLabel(iTempIndex).ZOrder
            iTempIndex = btnSchLabel(iTempIndex).HelpContextID
        Wend
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub setSchLabelToolTipText(ByVal iIndex As Integer)
'------------------------------------------------
'���ܣ�����ԤԼ��ǩ��Zorder����ʾ����ǰ��
'������ iIndex -- ԤԼ��ǩ������
'���أ� ��
'------------------------------------------------
    Dim iPoolIndex As Integer
    
    On Error GoTo err
    
    iPoolIndex = Val(btnSchLabel(iIndex).tag)
    btnSchLabel(iIndex).ToolTipText = "  ��ţ�" & mSchLabelPool(iPoolIndex).lng��� & vbCrLf & "  ������" & mSchLabelPool(iPoolIndex).str���� _
            & vbCrLf & "  ҽ�����ݣ�" & mSchLabelPool(iPoolIndex).strҽ������ & vbCrLf & "  ��ʼʱ�䣺" & Format(mSchLabelPool(iPoolIndex).dtStartTime, "HH:MM") _
            & vbCrLf & "  ����ʱ�䣺" & Format(mSchLabelPool(iPoolIndex).dtEndTime, "HH:MM")
            
    mstrOrderInfo = btnSchLabel(iIndex).ToolTipText
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub setSchLabelSelectTag(iIndex As Integer)
'------------------------------------------------
'���ܣ�����ԤԼ��ǩ�ı�ѡ�б��
'������ iIndex -- ԤԼ��ǩ������
'���أ� ��
'------------------------------------------------
    Dim iTempIndex As Integer
    Dim i As Integer
    Dim btnLabel As CommandButton
    
    On Error GoTo err
    
    For Each btnLabel In btnSchLabel
        If btnLabel.Index <> 0 Then
            btnLabel.Font.Bold = False
            btnLabel.Font.Size = mlngFontSize
        End If
    Next
    
    btnSchLabel(iIndex).Font.Bold = True
    btnSchLabel(iIndex).Font.Size = mlngFontSize + 2
    
    If btnSchLabel(iIndex).HelpContextID <> 0 Then
        iTempIndex = btnSchLabel(iIndex).HelpContextID
        While btnSchLabel(iIndex).HelpContextID <> btnSchLabel(iTempIndex).HelpContextID
            btnSchLabel(iTempIndex).Font.Bold = True
            btnSchLabel(iTempIndex).Font.Size = mlngFontSize + 3
            iTempIndex = btnSchLabel(iTempIndex).HelpContextID
        Wend
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function CanResizeLabel(iIndex As Integer) As Boolean
'------------------------------------------------
'���ܣ��ж����ԤԼ��ǩ�Ƿ���Ըı��ȣ������һ���ǩ��ֻ�����޸����һ����ǩ�Ŀ��
'������ iIndex -- ԤԼ��ǩ������
'���أ� True -- �����޸ģ�False -- �����޸�
'------------------------------------------------
        
    On Error GoTo err
    
    If btnSchLabel(iIndex).HelpContextID = 0 Or (btnSchLabel(iIndex).HelpContextID = mSchLabelPool(btnSchLabel(iIndex).tag).lngBtnIndex) Then
        CanResizeLabel = True
    Else
        CanResizeLabel = False
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub RestoreLabelPos(iIndex As Integer)
'------------------------------------------------
'���ܣ��ָ���ǩԭ�ȵ�λ��
'������ iIndex -- ԤԼ��ǩ������
'���أ�
'------------------------------------------------
    Dim iFirstIndex As Integer
    Dim iPoolIndex As Integer
    
    On Error GoTo err
    iPoolIndex = btnSchLabel(iIndex).tag
    iFirstIndex = mSchLabelPool(iPoolIndex).lngBtnIndex
    
    '�ȼ������ʱ�䣬�ټ��㿪ʼʱ��
    'lngMinutes = DateDiff("n", mSchLabelPool(iPoolIndex).dtStartTime, mSchLabelPool(iPoolIndex).dtEndTime)
    
    mSchLabelPool(iPoolIndex).dtStartTime = Format(mSchLabelPool(iPoolIndex).dtStartTime, "YYYY-MM-DD") & " " & Format(GetTimeFromXY(mlngOriginLeft, mlngOriginTop), "HH:MM")
    '����ʱ�䣬��Ҫ���ݱ�ǩ��������
    mSchLabelPool(iPoolIndex).dtEndTime = DateAdd("n", mlngOriginMinute, mSchLabelPool(iPoolIndex).dtStartTime)
    Call PutSchLabel(iFirstIndex, iPoolIndex)
   
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SetColor(ByVal lngID As Long, ByVal blnClear As Boolean)
On Error GoTo errH
    '���з���������ɫ
    Dim i As Long, j As Long
    
    Dim lngColorTabWorkSel As Long
    Dim intR As Integer, intG As Integer, intB As Integer
    Dim intTMP As Integer
    Dim lngColorTabWork(1) As Long
    
    With vsfTime
        intR = (mlngColorTabWork And &HFF) Mod 256
        intG = ((mlngColorTabWork And &HFF00) \ &H100) Mod 256
        intB = ((mlngColorTabWork And &HFF0000) \ &H10000) Mod 256
        
        intR = intR - 50
        If intR < 1 Then intR = 100
        
        intG = intG - 50
        If intG < 1 Then intG = 100
        
        intB = intB - 50
        If intB < 1 Then intB = 100
        
        lngColorTabWorkSel = RGB(intR, intG, intB)
        lngColorTabWork(0) = mlngColorTabWork
        
        
        intR = (mlngColorTabWork And &HFF) Mod 256
        intG = ((mlngColorTabWork And &HFF00) \ &H100) Mod 256
        intB = ((mlngColorTabWork And &HFF0000) \ &H10000) Mod 256
        
        intR = intR + 40
        If intR > 255 Then intR = 165
        
        intG = intG + 40
        If intG > 255 Then intG = 165
        
        intB = intB + 40
        If intB > 255 Then intB = 165
        lngColorTabWork(1) = RGB(intR, intG, intB)
        
        For i = 1 To .Rows - 1
            For j = 1 To .Cols - 1
                If .Cell(flexcpData, i, j) = lngID Then
                    'ͻ����ʾ
                    If Not blnClear Then
                        .Cell(flexcpBackColor, i, j) = lngColorTabWorkSel
                    End If
                ElseIf .Cell(flexcpData, i, j) <> 0 Then
                    '��ԭ������ɫ
                    
                    If mColordict.Item(.Cell(flexcpData, i, j)) = True Then
                        .Cell(flexcpBackColor, i, j) = lngColorTabWork(0)
                    Else '
                        .Cell(flexcpBackColor, i, j) = lngColorTabWork(1)
                    End If
                    
                Else
                    'ʲô������
                End If
            Next
        Next
    End With

    Exit Sub
errH:
    MsgBox err.Description, vbOKOnly, "���ԤԼ��ʾ"
End Sub

Private Sub ShowMouseTime(lngX As Long, lngY As Long)
'δ�϶�ʱ��״̬�¸������λ����ʾ��ǰ��������
On Error GoTo errH
    Dim lngRow As Long, lngCol As Long, lngTimeProjectID As Long
        
    If lngX <> 0 And lngY <> 0 Then
        lngRow = GetRowsFromY(lngY)
        lngCol = GetColsFromX(lngX)
        lngTimeProjectID = vsfTime.Cell(flexcpData, lngRow, lngCol)
        
        If lngTimeProjectID = 0 Then
            Call SetColor(0, True)
        Else
            Call SetColor(lngTimeProjectID, False)
        End If
    End If
    Exit Sub
errH:
    MsgBox err.Description
End Sub

Private Sub SetMouseTimePro(ByVal lngX As Long, ByVal lngY As Long, ByVal IsMove As Boolean)
'���������list�е�X,Y,�͵�ǰ������ʱ���ǩ���Զ���ʱ���ǩ�ƶ������λ�ö�Ӧ��ʱ�䣬����Ѿ��Ų��£�������ʾ ��һ��ʱ���
    
    Dim dtStartTime As Date         '��ʼʱ��
    Dim dtEndTime As Date           '����ʱ��
    Dim blnTimeOK As Boolean 'ʱ���������
    Dim blnNeedTestTime As Boolean '��Ҫ��֤ʱ��
    Dim dtStartTimeSQL As Date         '��ʼʱ��
    Dim dtEndTimeSQL As Date           '����ʱ��
    Dim intPoolIndex As Integer
    Dim i As Integer, j As Integer
    Dim iFirstIndex As Integer
    Dim iBtnIndex As Integer
    Dim blnFind As Boolean
    
    Dim lngLeft As Long
    Dim lngTop As Long
    On Error GoTo err
    
    blnNeedTestTime = False
    If mlngOrderID = 0 Then Exit Sub
    For i = 1 To UBound(mSchLabelPool)
        If mSchLabelPool(i).lngҽ��ID = mlngOrderID Then
            iBtnIndex = mSchLabelPool(i).lngBtnIndex
            Exit For
        End If
    Next
    
    If btnSchLabel(iBtnIndex).HelpContextID <> 0 Then   '�Ǳ�ǩ��
        '�����ҵ������ǩ���еĵ�һ����ǩ����
        iFirstIndex = mSchLabelPool(btnSchLabel(iBtnIndex).tag).lngBtnIndex
    Else    '���Ǳ�ǩ�飬��ǰ�������ǵ�һ������
        iFirstIndex = iBtnIndex
    End If

    Call AdjustLabelPos(iFirstIndex)

    mlngPoolIndex = btnSchLabel(iFirstIndex).tag
    mSchLabelPool(btnSchLabel(iFirstIndex).tag).bln�ѱ��� = False
    Call setSchLabelToolTipText(iFirstIndex)
    
    If IsMove Then
        '���ݿؼ���ǰ�����λ�ü�������ʱ���
        lngLeft = btnSchLabel(iFirstIndex).Left
        lngTop = btnSchLabel(iFirstIndex).Top + 0.5 * (btnSchLabel(iFirstIndex).Height)
        dtStartTimeSQL = GetTimeFromXY(lngLeft, lngTop)
    Else
        '����X ,Y ��������ʱ���
        dtStartTimeSQL = GetTimeFromXY(lngX, lngY)
    End If
    dtEndTimeSQL = DateAdd("n", DateDiff("n", mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtStartTime, mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtEndTime), dtStartTimeSQL)
    intPoolIndex = btnSchLabel(iBtnIndex).tag
    
    For i = 1 To UBound(mSchTimeProject)
         If mSchTimeProject(i).dtStartTime <= dtStartTimeSQL And mSchTimeProject(i).dtEndTime > dtStartTimeSQL Then
            dtStartTimeSQL = mSchTimeProject(i).dtStartTime
            dtEndTimeSQL = mSchTimeProject(i).dtEndTime
            blnFind = True
            Exit For
         End If
    Next
    
    If Not blnFind Then
        Exit Sub
    End If
  
    blnNeedTestTime = True
    If UBound(mSchLabelPool) = 1 Then
        dtStartTime = dtStartTimeSQL
        dtEndTime = dtEndTimeSQL
    Else
    
        If mSchLabelPool(1).dtStartTime >= Format(mdtSchDate, "YYYY-MM-DD") & dtEndTimeSQL Then
            '���������˳�whileѭ��
            dtStartTime = dtStartTimeSQL
            dtEndTime = dtEndTimeSQL
            blnNeedTestTime = False
        End If
        
        '������Ȼ����ʱ��ο�ʼʱ����ΪԤԼ��ʼʱ����Ч
        blnTimeOK = True
        dtStartTime = dtStartTimeSQL
        dtEndTime = DateAdd("n", DateDiff("n", mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtStartTime, mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtEndTime), dtStartTime)
    
        For i = 1 To UBound(mSchLabelPool)
            '�ܿ���ǰʱ���
            If intPoolIndex <> i Then '��Ҫ�޸��������
                If Not ((Format(mSchLabelPool(i).dtStartTime, "hh:mm:ss") < dtStartTime And Format(mSchLabelPool(i).dtEndTime, "hh:mm:ss") <= dtStartTime) Or (Format(mSchLabelPool(i).dtStartTime, "hh:mm:ss") >= dtEndTime And Format(mSchLabelPool(i).dtEndTime, "hh:mm:ss") > dtEndTime)) Then
                    blnTimeOK = False
                    Exit For
                End If
            End If
        Next
        
        If Not blnTimeOK Then
        
            For i = 1 To UBound(mSchLabelPool)
                If blnNeedTestTime And Format(mSchLabelPool(i).dtStartTime, "hh:mm:ss") >= dtStartTimeSQL And Format(mSchLabelPool(i).dtStartTime, "hh:mm:ss") <= dtEndTimeSQL Then  'And iBtnIndex <> i
                    blnTimeOK = True
                    
                    If blnTimeOK And blnNeedTestTime Then
                        dtStartTime = Format(mSchLabelPool(i).dtEndTime, "hh:mm:ss") 'rsSchedule!ԤԼ����ʱ�� 'Format(rsSchedule!ԤԼ����ʱ��, "hh-mm-ss")
                        dtEndTime = DateAdd("n", DateDiff("n", mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtStartTime, mSchLabelPool(btnSchLabel(iFirstIndex).tag).dtEndTime), dtStartTime)
                    
                        For j = 1 To UBound(mSchLabelPool)
                            If intPoolIndex <> j Then '��Ҫ�޸��������
                                If Format(mSchLabelPool(i).dtStartTime, "hh:mm:ss") >= dtStartTimeSQL And Format(mSchLabelPool(i).dtStartTime, "hh:mm:ss") <= dtEndTimeSQL Then
                                    If Not ((Format(mSchLabelPool(j).dtStartTime, "hh:mm:ss") < dtStartTime And Format(mSchLabelPool(j).dtEndTime, "hh:mm:ss") <= dtStartTime) Or (Format(mSchLabelPool(j).dtStartTime, "hh:mm:ss") >= dtEndTime And Format(mSchLabelPool(j).dtEndTime, "hh:mm:ss") > dtEndTime)) Then
                                        blnTimeOK = False
                                    End If
                                End If
                            End If
                        Next
    
                        If blnTimeOK Then
                            blnNeedTestTime = False
                        End If
                    End If
                End If
            Next
        End If
    End If

    Call ModifySchPoolTime(btnSchLabel(iFirstIndex).tag, dtStartTime)
    '���°ڷű�ǩ
    Call PutSchLabel(iFirstIndex, btnSchLabel(iFirstIndex).tag)

    If iFirstIndex <> iBtnIndex Then
        mlngBaseX = lngX
    End If
    
    RaiseEvent OnSchLabelModifed(iFirstIndex)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
