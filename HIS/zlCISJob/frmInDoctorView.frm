VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmInDoctorView 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "סԺһ��"
   ClientHeight    =   8625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13440
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   13440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7065
      Left            =   0
      ScaleHeight     =   7065
      ScaleWidth      =   12000
      TabIndex        =   0
      Top             =   480
      Width           =   12000
   End
   Begin XtremeCommandBars.CommandBars cbsSub 
      Left            =   120
      Top             =   840
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmInDoctorView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'�¼�
Public Event ViewPACSImage(ByVal ҽ��ID As Long) 'Ҫ����й�Ƭ
Public Event ResizeForm(ByVal bytFunc As Long)  '���ý��� 1-�Ŵ�;0-��ԭ

Private WithEvents mmessageManager As TimeLineMessageManager     '��Ϣ������
Attribute mmessageManager.VB_VarHelpID = -1
Private WithEvents mtimeLineControl As TimeLineControl           'ʱ����ؼ�
Attribute mtimeLineControl.VB_VarHelpID = -1

'��������
Private Const M_CON_TYPE_NORMAL As String = "��ͨ�ı�"
Private Const M_CON_TYPE_LAYOUT As String = "�Ű��ı�"
Private Const M_CON_TYPE_GROUP As String = "����"
Private Const M_CON_TYPE_COUNT As String = "��������"
Private Const M_CON_TYPE_DIVISION As String = "�ָ�����"
Private Const M_CON_TYPE_CONTINUOUS As String = "��������"
Private Const M_CON_TYPE_TICK As String = "ʱ������"
Private Const M_CON_TYPE_MEASURE As String = "�������"
Private Const M_CON_TYPE_MEASUREVERTICALTEXT As String = "�����ı�"
Private Const M_CON_TYPE_CUSTOMTICK As String = "�Զ���ʱ��"
Private Const M_CON_TYPE_DATAAREA As String = "��������"

'��������
Private Const M_CON_KEY_��ҩ����_�� As String = "K_��ҩ����_��"
Private Const M_CON_KEY_��ҩ����_�� As String = "K_��ҩ����_��"
Private Const M_CON_KEY_���� As String = "K_�����ļ�"
Private Const M_CON_KEY_��� As String = "K_���"
Private Const M_CON_KEY_���� As String = "K_����"
Private Const M_CON_KEY_�������� As String = "K_��������"
Private Const M_CON_KEY_�������� As String = "K_��������"
Private Const M_CON_KEY_���� As String = "K_����"
Private Const M_CON_KEY_סԺ���� As String = "K_סԺ����"
Private Const M_CON_KEY_���������� As String = "K_����������"
Private Const M_CON_KEY_�����ı� As String = "K_�����ı�"

'��ɫ���:ǳ��,����,�ʺ�,���,����   ������һ������һ��
Private Enum CONST_COLOR
    '��ɫ
    COLOR_ǳ�� = &HC0C0FF          'ǳ��
    COLOR_���� = &H8080FF
    COLOR_�ʺ� = &HFF&
    COLOR_��� = &HC0&
    COLOR_���� = &H80&
    '��ɫ
    COLOR_ǳ�� = &HC0E0FF
    COLOR_���� = &H80C0FF
    COLOR_�ʳ� = &H80FF&
    COLOR_��� = &H40C0&
    COLOR_���� = &H4080&
    '��ɫ
    COLOR_ǳ�� = &HC0FFFF
    COLOR_���� = &H80FFFF
    COLOR_�ʻ� = &HFFFF&
    COLOR_��� = &HC0C0&
    COLOR_���� = &H8080&
    '��ɫ
    COLOR_ǳ�� = &HC0FFC0
    COLOR_���� = &H80FF80
    COLOR_���� = &HFF00&
    COLOR_���� = &HC000&
    COLOR_���� = &H8000&
    '��ɫ
    COLOR_ǳ�� = &HFFFFC0
    COLOR_���� = &HFFFF80
    COLOR_���� = &HFFFF00
    COLOR_���� = &HC0C000
    COLOR_���� = &H808000
    '��ɫ
    COLOR_ǳ�� = &HFFC0C0
    COLOR_���� = &HFF8080
    COLOR_���� = &HFF0000
    COLOR_���� = &HC00000
    COLOR_���� = &H800000
    '��ɫ
    COLOR_ǳ�� = &HFFC0FF
    COLOR_���� = &HFF80FF
    COLOR_���� = &HFF00FF
    COLOR_���� = &HC000C0
    COLOR_���� = &H800080
    '��ɫ
    COLOR_��ɫ = &H80000005
    COLOR_FORMBK = &H8000000B
    COLOR_CENTERBK = &H808000
End Enum

'��갴��
Private Enum CONST_MouseButtons
    MouseButtons_Left = &H100000
    MouseButtons_None = 0
    MouseButtons_Middle = &H400000
    MouseButtons_Right = &H200000
    MouseButtons_XButton1 = &H800000
    MouseButtons_XButton2 = &H1000000
End Enum

Private mlng����ID As Long
Private mlng��ҳID As Long
Private mlng����ID As Long
Private mintBaby As Integer

Private mDatBegin As Date      'ÿһҳ�Ŀ�ʼ����
Private mDatEnd As Date        'ÿһҳ�Ľ�������
Private mDatIn As Date       '��¼��Ժ����
Private mdatOut As Date      '��¼��Ժ����
Private mudtTimeLine As TimeLineData     '
Private mudtDesign As TimeLineDesignInfo
Private mbytFont As Byte           '0-С����;1-������ ��ʾ������
Private mlngҽ��ID As Long
Private mlngӦ�÷�ʽ  As Long         '0-����;1-����ʹ��;2-��������������
Private mrsFrequency   As ADODB.Recordset     '������ĿƵ��   Ƶ��, ���, ��ʼ, ����, ���
Private mrs����ʱ�� As ADODB.Recordset         '�������ʱ��    ��ʼ, ����, ���
Private mrs������Ŀ As ADODB.Recordset          '���������Ŀ

Private mobjPopup As CommandBarPopup     '��ҳ��ť
Private mlngDay  As Long               '�������(һҳ)ȡֵ��ΧĬ�� 7��
Private mlngPages As Long              '��ҳ��

'��������
Private mstr���¿�ʼʱ�� As String    '������ =5 ģ��� 1255 ���Ʊ�׼���µ�(����ר��)ÿ��6��ʱ�����״ο�ʼ��ʱ�㣬6��ʱ����Ϊ4Сʱ���磺����ֵΪ4��6��ʱ��ֱ�Ϊ��4,8,12,16,20,24
Private mstr��ʾ���� As String   '������=42 ģ���=1255  ���ܲ�����ʾ��������
Private mblnMeasureArea As Boolean   'T-��ʾ��������,F-���ػ�������


'-----------------------------------------------------------------------------------------------------------------------------
'ˢ�½ӿ�
Public Function zlRefresh(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng����ID As Long, ByVal intBaby As Integer) As Boolean
'���ܣ�
'����:objFrmMain-������
'lngPatiID-����ID
'lngMainID-��ҳID
'lng����ID-����ID
'intBaby-Ӥ�����
    
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mlng����ID = lng����ID
    mintBaby = IIf(intBaby > 0, intBaby, 0)
    Call FuncLoadPages
End Function

Private Sub cbsSub_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim arrPara As Variant
    Dim lngPage As Long
    Dim i   As Long
    Dim udtDataItem As DataItem
    
    Select Case Control.ID
    
    Case conMenu_View_Jump
        arrPara = Split(Control.Parameter, ",")
        lngPage = Val(arrPara(0))
        mDatBegin = CDate(arrPara(1))
        mDatEnd = CDate(arrPara(2))
    Case conMenu_View_Forward, conMenu_View_Backward   '��һҳ,��һҳ
        arrPara = Split(mobjPopup.Parameter, ",")
        If Control.ID = conMenu_View_Forward Then
            lngPage = Val(arrPara(0)) - 1
            mDatBegin = mDatBegin - mlngDay
            mDatEnd = mDatEnd - mlngDay
        ElseIf Control.ID = conMenu_View_Backward Then
            lngPage = Val(arrPara(0)) + 1
            mDatBegin = mDatBegin + mlngDay
            mDatEnd = mDatEnd + mlngDay
        End If
    Case conMenu_Img_Look  '��Ƭ
        If mlngҽ��ID <> 0 Then
            RaiseEvent ViewPACSImage(mlngҽ��ID)
        End If
    Case conMenu_View_Show
        If Not mblnMeasureArea Then
            Control.Caption = "���ػ���"
            Control.ToolTipText = "���ػ���"
            Control.IconId = conMenu_Manage_Up
        Else
            Control.Caption = "��ʾ����"
            Control.ToolTipText = "��ʾ����"
            Control.IconId = conMenu_Manage_Down
        End If
        mblnMeasureArea = Not mblnMeasureArea
        Call SetFontSize(mbytFont)
    Case conMenu_Tool_Assistant
        Control.Checked = Not Control.Checked
        mtimeLineControl.IsShowReticle = Control.Checked
    Case conMenu_View_Navigatebeginning
        mtimeLineControl.ScrollToLeft
    Case conMenu_View_Navigateend
        mtimeLineControl.ScrollToRight
    Case conMenu_Process_Zoom
        Control.Checked = Not Control.Checked
        If Control.Checked Then
            Control.IconId = conMenu_Process_Small
            RaiseEvent ResizeForm(1)
            Control.Caption = "��С"
            Control.ToolTipText = "��С"
        Else
            Control.IconId = conMenu_Process_Zoom
            RaiseEvent ResizeForm(0)
            Control.Caption = "�Ŵ�"
            Control.ToolTipText = "�Ŵ�"
        End If
    End Select
    
    Select Case Control.ID
    Case conMenu_View_Jump, conMenu_View_Forward, conMenu_View_Backward
        mobjPopup.Caption = "��" & lngPage & "ҳ��" & Format(mDatBegin, "YYYY-MM-DD") & "��" & Format(mDatEnd, "YYYY-MM-DD")
        mobjPopup.Parameter = lngPage & "," & Format(mDatBegin, "YYYY-MM-DD") & "," & Format(mDatEnd, "YYYY-MM-DD")
        mobjPopup.SetFocus
        For i = 1 To mobjPopup.CommandBar.Controls.Count
            If i = lngPage Then
                mobjPopup.CommandBar.Controls(i).Checked = True
            Else
                mobjPopup.CommandBar.Controls(i).Checked = False
            End If
        Next
        Call FuncCreateTimeLine
    End Select
End Sub

'---------------------------------------------------------------------------------------------------------------------------------
Private Sub cbsSub_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call cbsSub.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    On Error Resume Next
    With picMain
        .Left = lngLeft: .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = lngBottom - .Top
    End With
    mtimeLineControl.SetParentControl picMain.hwnd
End Sub

Private Sub cbsSub_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim arrPara As Variant
    
    arrPara = Split(mobjPopup.Parameter, ",")
    mobjPopup.Visible = mlngPages > 1
    
    Select Case Control.ID
    Case conMenu_View_Forward, conMenu_View_Backward   '��һҳ,��һҳ
        Control.Visible = mlngPages > 1
        If Control.Visible Then
            If Val(arrPara(0)) = 1 And Control.ID = conMenu_View_Forward Then
                Control.Enabled = False
            ElseIf Val(arrPara(0)) = mlngPages And Control.ID = conMenu_View_Backward Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        End If
    Case conMenu_View_Navigatebeginning, conMenu_View_Navigateend
        Control.Visible = (mDatEnd - mDatBegin) > 15
    End Select
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim tlbContext As TimeLineBusinessContext     'ʱ����ҵ�񻷾�
    Dim strTmp As String
    
    '������ȡ
    mstr���¿�ʼʱ�� = zlDatabase.GetPara("���¿�ʼʱ��", glngSys, p�����¼����, "4")
    mstr��ʾ���� = zlDatabase.GetPara("���ܲ�����ʾ��������", glngSys, p�����¼����, "0")
    
    Set tlbContext = New TimeLineBusinessContext
    Set mmessageManager = tlbContext.MessageManager

    Set mtimeLineControl = New TimeLineControl
    Set mtimeLineControl.BusinessContext = tlbContext
    mtimeLineControl.DockMode = TimeLineDockStyle_Fill '��������
    strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\��ʾ����", "��ʾ����", "1")
    mblnMeasureArea = IIf(strTmp = "1", True, False)
    
    Call InitCommandBar
    
    strSQL = "Select Ƶ��, ���, ��ʼ, ����, ��� From ������ĿƵ��"
    Set mrsFrequency = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    strSQL = "Select ��ʼ, ����, ��� From �������ʱ�� Where ���� = 1"
    Set mrs����ʱ�� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    strSQL = "Select ���, NVL(�����,0) as ����� From ���������Ŀ"
    Set mrs������Ŀ = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)

End Sub

Private Sub Form_Resize()
    cbsSub_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.Visible Then
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\��ʾ����", "��ʾ����", IIf(mblnMeasureArea, "1", "0"))
    End If
    Set mtimeLineControl = Nothing
    Set mmessageManager = Nothing
End Sub

Private Function FuncMakeXMLDesign(ByRef udtDesign As TimeLineDesignInfo) As String
    Dim strDesign As String
    Dim strTmp As String
    Dim intDisplayVal As Integer, intTickStartTime As Integer
    Dim intTemp As Integer
    Dim i As Long

    Dim colTemp As Collection
    Dim udtTick  As DesignInfoTickRange
    
    strDesign = "<?xml version=""1.0"" encoding=""utf-16""?>" & vbCrLf & _
        "<TimeLineDesignInfo xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">" & vbCrLf
    With udtDesign
        Set colTemp = mudtTimeLine.colMeasureData
        If mbytFont = 1 Then
            .RowHeight = 22
            If colTemp.Count > 0 And colTemp.Count <= 3 Then
                .MeasureTitleWidth = 180 \ colTemp.Count
            Else
                .MeasureTitleWidth = 65
            End If
            .GridMinHeight = 220
            .DateTitleFont = Replace(.DateTitleFont, "9pt", "12pt")
            .TickTitleFont = Replace(.TickTitleFont, "9pt", "12pt")
            .Height = 200
            If mblnMeasureArea Then
                .TickWidth = 25
                .ShowTick = True
            Else
                .TickWidth = 150
                .ShowTick = False
            End If
        Else
            .RowHeight = 17
            If colTemp.Count > 0 And colTemp.Count <= 3 Then
                .MeasureTitleWidth = 150 \ colTemp.Count
            Else
                .MeasureTitleWidth = 50
            End If

            .GridMinHeight = 220
            .DateTitleFont = Replace(.DateTitleFont, "12pt", "9pt")
            .TickTitleFont = Replace(.TickTitleFont, "12pt", "9pt")
            .Height = 160
            If mblnMeasureArea Then
                .TickWidth = 20
                .ShowTick = True
            Else
                .TickWidth = 120
                .ShowTick = False
            End If
        End If

        strDesign = strDesign & _
                    IIf(.BackgroundColor <> Empty, Space(2) & "<BackgroundColor>" & .BackgroundColor & "</BackgroundColor>" & vbCrLf, "") & _
                    IIf(.DateTitle <> Empty, Space(2) & "<DateTitle>" & .DateTitle & "</DateTitle>" & vbCrLf, "") & _
                    IIf(.DateTitleFont <> Empty, Space(2) & "<DateTitleFont>" & .DateTitleFont & "</DateTitleFont>" & vbCrLf, "") & _
                    IIf(.DateTitleColor <> Empty, Space(2) & "<DateTitleColor>" & .DateTitleColor & "</DateTitleColor>" & vbCrLf, "") & _
                    IIf(.DateStart <> Empty, Space(2) & "<DateStart>" & .DateStart & "</DateStart>" & vbCrLf, "") & _
                    IIf(.dateEnd <> Empty, Space(2) & "<DateEnd>" & .dateEnd & "</DateEnd>" & vbCrLf, "") & _
                    IIf(.ShowTick <> Empty, Space(2) & "<ShowTick>" & IIf(.ShowTick, "true", "false") & "</ShowTick>" & vbCrLf, "") & _
                    IIf(.ShowFullDate <> Empty, Space(2) & "<ShowFullDate>" & IIf(.ShowFullDate, "true", "false") & "</ShowFullDate>" & vbCrLf, "") & _
                    IIf(.TickTitle <> Empty, Space(2) & "<TickTitle>" & .TickTitle & "</TickTitle>" & vbCrLf, "") & _
                    IIf(.TickTitleFont <> Empty, Space(2) & "<TickTitleFont>" & .TickTitleFont & "</TickTitleFont>" & vbCrLf, "") & _
                    IIf(.TickTitleColor <> Empty, Space(2) & "<TickTitleColor>" & .TickTitleColor & "</TickTitleColor>" & vbCrLf, "") & _
                    IIf(.TickWidth <> Empty, Space(2) & "<TickWidth>" & .TickWidth & "</TickWidth>" & vbCrLf, "")
        strDesign = strDesign & Space(2) & "<DesignInfoTickRangeList>" & vbCrLf
        strTmp = ""
        If .TickRangeListCount = 0 Then .TickRangeListCount = TICK_6   'ȱʡ���Ϊ6
        intTemp = 24 \ .TickRangeListCount
        intDisplayVal = Val(mstr���¿�ʼʱ��)
        intTickStartTime = 0
        For i = 1 To .TickRangeListCount
            strTmp = strTmp & _
                Space(4) & "<DesignInfoTickRange>" & vbCrLf & _
                    Space(6) & "<DisplayValue>" & intDisplayVal & "</DisplayValue>" & vbCrLf & _
                    Space(6) & "<TickStartTime>" & intTickStartTime & ":0" & "</TickStartTime>" & vbCrLf & _
                Space(4) & "</DesignInfoTickRange>" & vbCrLf
            intDisplayVal = intDisplayVal + intTemp  '��һ����ʾֵ
            intTickStartTime = intTickStartTime + intTemp
        Next
    
        strDesign = strDesign & strTmp & "</DesignInfoTickRangeList>" & vbCrLf
        '�������
        If mbytFont = 1 Then
            If .DateFont <> Empty Then .DateFont = Replace(.DateFont, "9pt", "12pt")
            If .TickFont <> Empty Then .TickFont = Replace(.TickFont, "9pt", "12pt")
        Else
            If .DateFont <> Empty Then .DateFont = Replace(.DateFont, "12pt", "9pt")
            If .TickFont <> Empty Then .TickFont = Replace(.TickFont, "12pt", "9pt")
        End If
        strDesign = strDesign & _
                    IIf(.DateFont <> Empty, Space(2) & "<DateFont>" & .DateFont & "</DateFont>" & vbCrLf, "") & _
                    IIf(.TickFont <> Empty, Space(2) & "<TickFont>" & .TickFont & "</TickFont>" & vbCrLf, "") & _
                    IIf(.MergePeriodWidth <> Empty, Space(2) & "<MergePeriodWidth>" & .MergePeriodWidth & "</MergePeriodWidth>" & vbCrLf, "") & _
                    IIf(.EmptyDataMergeDayCount <> Empty, Space(2) & "<EmptyDataMergeDayCount>" & .EmptyDataMergeDayCount & "</EmptyDataMergeDayCount>" & vbCrLf, "") & _
                    IIf(.EmptyDataMergePeriodWidth <> Empty, Space(2) & "<EmptyDataMergePeriodWidth>" & .EmptyDataMergePeriodWidth & "</EmptyDataMergePeriodWidth>" & vbCrLf, "") & _
                    IIf(.PaddingLeft <> Empty, Space(2) & "<PaddingLeft>" & .PaddingLeft & "</PaddingLeft>" & vbCrLf, "") & _
                    IIf(.PaddingTop <> Empty, Space(2) & "<PaddingTop>" & .PaddingTop & "</PaddingTop>" & vbCrLf, "") & _
                    IIf(.PaddingRight <> Empty, Space(2) & "<PaddingRight>" & .PaddingRight & "</PaddingRight>" & vbCrLf, "") & _
                    IIf(.PaddingBottom <> Empty, Space(2) & "<PaddingBottom>" & .PaddingBottom & "</PaddingBottom>" & vbCrLf, "") & _
                    IIf(.RowHeight <> Empty, Space(2) & "<RowHeight>" & .RowHeight & "</RowHeight>" & vbCrLf, "")
    
        strDesign = strDesign & _
                    Space(2) & "<Measure>" & vbCrLf & _
                        IIf(.MeasureTitleWidth <> Empty, Space(4) & "<MeasureTitleWidth>" & .MeasureTitleWidth & "</MeasureTitleWidth>" & vbCrLf, "") & _
                        IIf(.GridMinHeight <> Empty, Space(4) & "<GridMinHeight>" & .GridMinHeight & "</GridMinHeight>" & vbCrLf, "") & _
                        IIf(.TopFixedSmallRowCount <> Empty, Space(4) & "<TopFixedSmallRowCount>" & .TopFixedSmallRowCount & "</TopFixedSmallRowCount>" & vbCrLf, "") & _
                        IIf(.BottomFixedSmallRowCount <> Empty, Space(4) & "<BottomFixedSmallRowCount>" & .BottomFixedSmallRowCount & "</BottomFixedSmallRowCount>" & vbCrLf, "") & _
                        IIf(.GridYSplitCount <> Empty, Space(4) & "<GridYSplitCount>" & .GridYSplitCount & "</GridYSplitCount>" & vbCrLf, "") & _
                        IIf(.GridYSmallSplitCount <> Empty, Space(4) & "<GridYSmallSplitCount>" & .GridYSmallSplitCount & "</GridYSmallSplitCount>" & vbCrLf, "") & _
                        IIf(.Height <> Empty, Space(4) & "<Height>" & .Height & "</Height>" & vbCrLf, "") & _
                    Space(2) & "</Measure>" & vbCrLf

    End With
    strDesign = strDesign & "</TimeLineDesignInfo>"
    
    FuncMakeXMLDesign = strDesign
End Function

Private Function FuncMakeXMLDataList(ByRef colList As Collection, Optional ByVal intSpace As Integer) As String
'����:
'colList-DataItem\DataInfo�ļ���
    Dim udtItem As DataItem
    Dim udtInfo As DataInfo
    Dim varItem As Variant
    Dim strRet As String
    Dim strHead As String
    Dim strFoot As String
    
    If colList Is Nothing Then Exit Function

    If colList.Count = 0 Then Exit Function

    If TypeName(colList(colList.Count)) = TypeName(udtItem) Then
        strHead = Space(intSpace) & "<ListDataItem>" & vbCrLf
        strFoot = Space(intSpace) & "</ListDataItem>" & vbCrLf
    ElseIf TypeName(colList(colList.Count)) = TypeName(udtInfo) Then
        strHead = Space(intSpace) & "<ListDataInfo>" & vbCrLf
        strFoot = Space(intSpace) & "</ListDataInfo>" & vbCrLf
    End If
    For Each varItem In colList
        strRet = strRet & FuncMakeXMLData(varItem, (intSpace + 2))
    Next
    FuncMakeXMLDataList = strHead & strRet & strFoot
End Function

Private Function FuncMakeXMLData(ByRef varData As Variant, Optional ByVal intSpace As Integer) As String
    Dim strRet As String
    Dim udtItem As DataItem
    Dim udtInfo As DataInfo

    If TypeName(varData) = TypeName(udtItem) Then
        udtItem = varData
        With udtItem
            If mbytFont = 1 Then
                If .Font <> Empty Then .Font = Replace(.Font, "9pt", "12pt")
                If .HotspotFont <> Empty Then .HotspotFont = Replace(.HotspotFont, "9pt", "12pt")
                If .TitleFont <> Empty Then .TitleFont = Replace(.TitleFont, "9pt", "12pt")
            Else
                If .Font <> Empty Then .Font = Replace(.Font, "12pt", "9pt")
                If .HotspotFont <> Empty Then .HotspotFont = Replace(.HotspotFont, "12pt", "9pt")
                If .TitleFont <> Empty Then .TitleFont = Replace(.TitleFont, "12pt", "9pt")
            End If
            strRet = Space(intSpace + 2) & "<DataItem>" & vbCrLf
                '��������
                strRet = strRet & IIf(.GraphType <> Empty, Space(intSpace + 4) & "<GraphType>" & .GraphType & "</GraphType>" & vbCrLf, "")
                strRet = strRet & IIf(.Title <> Empty, Space(intSpace + 4) & "<Title>" & .Title & "</Title>" & vbCrLf, "")
                strRet = strRet & IIf(.TextColor <> Empty, Space(intSpace + 4) & "<TextColor>" & .TextColor & "</TextColor>" & vbCrLf, "")
                strRet = strRet & IIf(.Color <> Empty, Space(intSpace + 4) & "<Color>" & .Color & "</Color>" & vbCrLf, "")
                strRet = strRet & IIf(.Font <> Empty, Space(intSpace + 4) & "<Font>" & .Font & "</Font>" & vbCrLf, "")
                strRet = strRet & IIf(.BackgroundColor <> Empty, Space(intSpace + 4) & "<BackgroundColor>" & .BackgroundColor & "</BackgroundColor>" & vbCrLf, "")
                strRet = strRet & IIf(.ShowHotspotEffect = True, Space(intSpace + 4) & "<ShowHotspotEffect>true</ShowHotspotEffect>" & vbCrLf, "")
                strRet = strRet & IIf(.HotspotFont <> Empty, Space(intSpace + 4) & "<HotspotFont>" & .HotspotFont & "</HotspotFont>" & vbCrLf, "")
                strRet = strRet & IIf(.HotspotColor <> Empty, Space(intSpace + 4) & "<HotspotColor>" & .HotspotColor & "</HotspotColor>" & vbCrLf, "")
                strRet = strRet & IIf(.ShowHotspotCursor = True, Space(intSpace + 4) & "<ShowHotspotCursor>true</ShowHotspotCursor>" & vbCrLf, "")
                strRet = strRet & IIf(.TitleFont <> Empty, Space(intSpace + 4) & "<TitleFont>" & .TitleFont & "</TitleFont>" & vbCrLf, "")
                strRet = strRet & IIf(.TitleColor <> Empty, Space(intSpace + 4) & "<TitleColor>" & .TitleColor & "</TitleColor>" & vbCrLf, "")
                strRet = strRet & IIf(.GroupPosition <> Empty, Space(intSpace + 4) & "<GroupPosition>" & .GroupPosition & "</GroupPosition>" & vbCrLf, "")
                '˽������
                Select Case .GraphType
                
                Case M_CON_TYPE_GROUP   '����
                    'BackgroundColor,Title,GraphType,GroupPosition
                Case M_CON_TYPE_COUNT, M_CON_TYPE_CONTINUOUS, M_CON_TYPE_NORMAL, M_CON_TYPE_TICK, M_CON_TYPE_LAYOUT '��������,��������,��ͨ�ı�,ʱ������,�Ű��ı�
                    'GraphType,Title,BackgroundColor,TextColor
                    If .GraphType = M_CON_TYPE_LAYOUT Then
                        strRet = strRet & IIf(.BorderColor <> Empty, Space(intSpace + 4) & "<BorderColor>" & .BorderColor & "</BorderColor>" & vbCrLf, "")
                    ElseIf .GraphType = M_CON_TYPE_CONTINUOUS Then
                        strRet = strRet & IIf(.Effect <> Empty, Space(intSpace + 4) & "<Effect>" & .Effect & "</Effect>" & vbCrLf, "")
                    End If
                Case M_CON_TYPE_DIVISION '�ָ�����
                    'Font,TextColor,BackgroundColor,Title
                     strRet = strRet & IIf(.SplitString <> Empty, Space(intSpace + 4) & "<SplitString>" & .SplitString & "</SplitString>" & vbCrLf, "") & _
                    IIf(.SplitCount <> Empty, Space(intSpace + 4) & "<SplitCount>" & .SplitCount & "</SplitCount>" & vbCrLf, "")
                Case M_CON_TYPE_MEASURE      '�������
                    'Color,Title
                    strRet = strRet & IIf(.Unit <> Empty, Space(intSpace + 4) & "<Unit>" & .Unit & "</Unit>" & vbCrLf, "") & _
                    IIf(.MinValue <> Empty, Space(intSpace + 4) & "<MinValue>" & .MinValue & "</MinValue>" & vbCrLf, "") & _
                    IIf(.MaxValue <> Empty, Space(intSpace + 4) & "<MaxValue>" & .MaxValue & "</MaxValue>" & vbCrLf, "") & _
                    IIf(.SplitNum <> Empty, Space(intSpace + 4) & "<SplitNum>" & .SplitNum & "</SplitNum>" & vbCrLf, "") & _
                    IIf(.SplitScale <> Empty, Space(intSpace + 4) & "<SplitScale>" & .SplitScale & "</SplitScale>" & vbCrLf, "") & _
                    IIf(.IsDataDynamicExpansion <> Empty, Space(intSpace + 4) & "<IsDataDynamicExpansion>" & IIf(.IsDataDynamicExpansion, "true", "false") & "</IsDataDynamicExpansion>" & vbCrLf, "") & _
                    IIf(.ShadowTitle <> Empty, Space(intSpace + 4) & "<ShadowTitle>" & .ShadowTitle & "</ShadowTitle>" & vbCrLf, "") & _
                    IIf(.ShadowColor <> Empty, Space(intSpace + 4) & "<ShadowColor>" & .ShadowColor & "</ShadowColor>" & vbCrLf, "") & _
                    IIf(.BalloonColor <> Empty, Space(intSpace + 4) & "<BalloonColor>" & .BalloonColor & "</BalloonColor>" & vbCrLf, "") & _
                    IIf(.BalloonTitle <> Empty, Space(intSpace + 4) & "<BalloonTitle>" & .BalloonTitle & "</BalloonTitle>" & vbCrLf, "") & _
                    IIf(.LegendType <> Empty, Space(intSpace + 4) & "<LegendType>" & .LegendType & "</LegendType>" & vbCrLf, "") & _
                    IIf(.ShadowLegendType <> Empty, Space(intSpace + 4) & "<ShadowLegendType>" & .ShadowLegendType & "</ShadowLegendType>" & vbCrLf, "") & _
                    IIf(.BalloonLegendType <> Empty, Space(intSpace + 4) & "<BalloonLegendType>" & .BalloonLegendType & "</BalloonLegendType>" & vbCrLf, "")
                Case M_CON_TYPE_MEASUREVERTICALTEXT    '�����ı�

                Case M_CON_TYPE_CUSTOMTICK   '�Զ���ʱ��
                    strRet = strRet & IIf(.StartDate <> Empty, Space(intSpace + 4) & "<StartDate>" & .StartDate & "</StartDate>" & vbCrLf, "") & _
                    IIf(.EndDate <> Empty, Space(intSpace + 4) & "<EndDate>" & .EndDate & "</EndDate>" & vbCrLf, "") & _
                    IIf(.FixedTick <> Empty, Space(intSpace + 4) & "<FixedTick>" & .FixedTick & "</FixedTick>" & vbCrLf, "") & _
                    IIf(.EquantTick <> Empty, Space(intSpace + 4) & "<EquantTick>" & .EquantTick & "</EquantTick>" & vbCrLf, "") & _
                    IIf(.EquantTickUnit <> Empty, Space(intSpace + 4) & "<EquantTickUnit>" & .EquantTickUnit & "</EquantTickUnit>" & vbCrLf, "") & _
                    IIf(.TickWidth <> Empty, Space(intSpace + 4) & "<TickWidth>" & .TickWidth & "</TickWidth>" & vbCrLf, "")
                Case M_CON_TYPE_DATAAREA     '��������
                    strRet = strRet & IIf(.LineColor <> Empty, Space(intSpace + 4) & "<LineColor>" & .LineColor & "</LineColor>" & vbCrLf, "") & _
                    IIf(.StartDate <> Empty, Space(intSpace + 4) & "<StartDate>" & .StartDate & "</StartDate>" & vbCrLf, "") & _
                    IIf(.EndDate <> Empty, Space(intSpace + 4) & "<EndDate>" & .EndDate & "</EndDate>" & vbCrLf, "") & _
                    IIf(.IsCollapse <> Empty, Space(intSpace + 4) & "<IsCollapse>" & .IsCollapse & "</IsCollapse>" & vbCrLf, "")
                End Select
                
                If Not .ListData Is Nothing Then
                    If .ListData.Count > 0 Then
                        strRet = strRet & FuncMakeXMLDataList(.ListData, intSpace + 4)
                    End If
                End If
                
            strRet = strRet & Space(intSpace + 2) & "</DataItem>" & vbCrLf
    
        End With
    ElseIf TypeName(varData) = TypeName(udtInfo) Then
        udtInfo = varData
        With udtInfo
            If mbytFont = 1 Then
                If .Font <> Empty Then .Font = Replace(.Font, "9pt", "12pt")
                If .HotspotFont <> Empty Then .HotspotFont = Replace(.HotspotFont, "9pt", "12pt")
            Else
                If .Font <> Empty Then .Font = Replace(.Font, "12pt", "9pt")
                If .HotspotFont <> Empty Then .HotspotFont = Replace(.HotspotFont, "12pt", "9pt")
            End If
            strRet = Space(intSpace + 2) & "<DataInfo>" & vbCrLf
            strRet = strRet & IIf(.Value <> Empty, Space(intSpace + 4) & "<Value>" & .Value & "</Value>" & vbCrLf, "")
            strRet = strRet & IIf(.Time <> Empty, Space(intSpace + 4) & "<Time>" & .Time & "</Time>" & vbCrLf, "")
            strRet = strRet & IIf(.RowNumber <> Empty, Space(intSpace + 4) & "<RowNumber>" & .RowNumber & "</RowNumber>" & vbCrLf, "")
            strRet = strRet & IIf(.TimeEnd <> Empty, Space(intSpace + 4) & "<TimeEnd>" & .TimeEnd & "</TimeEnd>" & vbCrLf, "")
            strRet = strRet & IIf(.Tag <> Empty, Space(intSpace + 4) & "<Tag>" & .Tag & "</Tag>" & vbCrLf, "")
            strRet = strRet & IIf(.BackgroundColor <> Empty, Space(intSpace + 4) & "<BackgroundColor>" & .BackgroundColor & "</BackgroundColor>" & vbCrLf, "")
            strRet = strRet & IIf(.TextColor <> Empty, Space(intSpace + 4) & "<TextColor>" & .TextColor & "</TextColor>" & vbCrLf, "")
            strRet = strRet & IIf(.Font <> Empty, Space(intSpace + 4) & "<Font>" & .Font & "</Font>" & vbCrLf, "")
            strRet = strRet & IIf(.HotspotFont <> Empty, Space(intSpace + 4) & "<HotspotFont>" & .HotspotFont & "</HotspotFont>" & vbCrLf, "")
            strRet = strRet & IIf(.ShowHotspotCursor = True, Space(intSpace + 4) & "<ShowHotspotCursor>true</ShowHotspotCursor>" & vbCrLf, "")
            strRet = strRet & IIf(.HotspotColor <> Empty, Space(intSpace + 4) & "<HotspotColor>" & .HotspotColor & "</HotspotColor>" & vbCrLf, "")
            strRet = strRet & IIf(.RowIndex <> Empty, Space(intSpace + 4) & "<RowIndex>" & .RowIndex & "</RowIndex>" & vbCrLf, "")
            strRet = strRet & IIf(.LegendType <> Empty, Space(intSpace + 4) & "<LegendType>" & .LegendType & "</LegendType>" & vbCrLf, "")
            strRet = strRet & IIf(.ShadowLegendType <> Empty, Space(intSpace + 4) & "<ShadowLegendType>" & .ShadowLegendType & "</ShadowLegendType>" & vbCrLf, "")
            strRet = strRet & IIf(.BalloonLegendType <> Empty, Space(intSpace + 4) & "<BalloonLegendType>" & .BalloonLegendType & "</BalloonLegendType>" & vbCrLf, "")
            strRet = strRet & IIf(.NumberValue <> Empty, Space(intSpace + 4) & "<NumberValue>" & .NumberValue & "</NumberValue>" & vbCrLf, "")
            strRet = strRet & IIf(.ShadowValue <> Empty, Space(intSpace + 4) & "<ShadowValue>" & .ShadowValue & "</ShadowValue>" & vbCrLf, "")
            strRet = strRet & IIf(.BalloonValue <> Empty, Space(intSpace + 4) & "<BalloonValue>" & .BalloonValue & "</BalloonValue>" & vbCrLf, "")
            strRet = strRet & IIf(.Tip <> Empty, Space(intSpace + 4) & "<Tip>" & .Tip & "</Tip>" & vbCrLf, "")
            strRet = strRet & IIf(.Group <> Empty, Space(intSpace + 4) & "<Group>" & .Group & "</Group>" & vbCrLf, "")
            strRet = strRet & Space(intSpace + 2) & "</DataInfo>" & vbCrLf
        End With
    End If
    FuncMakeXMLData = strRet
End Function

Private Function FuncMakeXMLTimeLine(udtData As TimeLineData) As String
    Dim strRet As String
    Dim strTmp As String
    Dim varItem As Variant
    Dim udtDataItem As DataItem
    Dim colTmp As Collection
    Dim objXML As zl9ComLib.clsXML
    Dim lngBegin As Long, lngEnd As Long

    strRet = "<?xml version=""1.0"" encoding=""utf-16""?>" & vbCrLf & _
            "<TimeLineData xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">" & vbCrLf
    
    '��ͷ����
    If Not udtData.colHeaderData Is Nothing Then
        strRet = strRet & _
            Space(2) & "<HeaderData>" & vbCrLf & _
            Space(4) & "<ListDataItem>" & vbCrLf
        For Each varItem In udtData.colHeaderData
            strRet = strRet & FuncMakeXMLData(varItem, 4)
        Next
        strRet = strRet & Space(4) & "</ListDataItem>" & vbCrLf & _
            Space(2) & "</HeaderData>" & vbCrLf
    End If
    'ҳ������
    If Not udtData.colFooterData Is Nothing Then
        strRet = strRet & _
            Space(2) & "<FooterData>" & vbCrLf & _
            Space(4) & "<ListDataItem>" & vbCrLf
        For Each varItem In udtData.colFooterData
            If varItem.Title = "������Ŀ" And mblnMeasureArea = False Then
                '���ػ�����Ŀ
            Else
                strRet = strRet & FuncMakeXMLData(varItem, 4)
            End If
            
        Next
        strRet = strRet & Space(4) & "</ListDataItem>" & vbCrLf & _
            Space(2) & "</FooterData>" & vbCrLf
    End If

    '�������
    If Not udtData.colMeasureData Is Nothing Then
        strRet = strRet & _
            Space(2) & "<MeasureData>" & vbCrLf & _
            Space(4) & "<ListDataItem>" & vbCrLf
        For Each varItem In udtData.colMeasureData
            strRet = strRet & FuncMakeXMLData(varItem, 4)
        Next
        strRet = strRet & Space(4) & "</ListDataItem>" & vbCrLf & _
            Space(2) & "</MeasureData>" & vbCrLf
    End If

    '�����ı�����
    If Not udtData.colMeasureVerticalText Is Nothing Then
        strRet = strRet & _
            Space(2) & "<MeasureVerticalText>" & vbCrLf & _
                Space(4) & "<ListDataItem>" & vbCrLf
        For Each varItem In udtData.colMeasureVerticalText
            strRet = strRet & FuncMakeXMLData(varItem, 4)
        Next
        strRet = strRet & Space(4) & "</ListDataItem>" & vbCrLf & _
            Space(2) & "</MeasureVerticalText>" & vbCrLf
    End If

    '�Զ���ʱ��
    If Not udtData.colCustomTick Is Nothing Then
        strRet = strRet & _
            Space(2) & "<CustomTick>" & vbCrLf & _
                Space(4) & "<ListDataItem>" & vbCrLf
        For Each varItem In udtData.colCustomTick
            strRet = strRet & FuncMakeXMLData(varItem, 4)
        Next
        strRet = strRet & Space(4) & "</ListDataItem>" & vbCrLf & _
            Space(2) & "</CustomTick>" & vbCrLf
    End If
    '��������
    If Not udtData.colDataArea Is Nothing Then
        strRet = strRet & _
            Space(2) & "<DataArea>" & vbCrLf & _
            Space(4) & "<ListDataItem>" & vbCrLf
        For Each varItem In udtData.colDataArea
            strRet = strRet & FuncMakeXMLData(varItem, 4)
        Next
        strRet = strRet & Space(4) & "</ListDataItem>" & vbCrLf & _
            Space(2) & "</DataArea>" & vbCrLf
    End If
    
    strRet = strRet & "</TimeLineData>"

    FuncMakeXMLTimeLine = strRet
End Function

Private Sub FuncClearUDT(varItem As Variant)
    Dim udtItem As DataItem
    Dim udtInfo As DataInfo
    
    Select Case TypeName(varItem)
    Case TypeName(udtItem)
        With varItem
            .BackgroundColor = Empty
            .BalloonLegendType = Empty
            .BalloonTitle = Empty
            .Color = Empty
            .EndDate = Empty
            .EquantTick = Empty
            .EquantTickUnit = Empty
            .FixedTick = Empty
            .Font = Empty
            .ShowHotspotEffect = False
            .HotspotFont = Empty
            .HotspotColor = Empty
            .GraphType = Empty
            .IsCollapse = Empty
            .IsDataDynamicExpansion = Empty
            .LegendType = Empty
            .LineColor = Empty
            Set .ListData = Nothing
            .MaxValue = Empty
            .MinValue = Empty
            .ShadowLegendType = Empty
            .ShadowTitle = Empty
            .SplitCount = Empty
            .SplitNum = Empty
            .SplitScale = Empty
            .SplitString = Empty
            .StartDate = Empty
            .TextColor = Empty
            .TickWidth = Empty
            .Title = Empty
            .Unit = Empty
            .BorderColor = Empty
            .ShowHotspotCursor = False
            .TitleColor = Empty
            .TitleFont = Empty
            .ShadowColor = Empty
            .BalloonColor = Empty
            .GroupPosition = Empty
            .Effect = Empty
            
            '
            .ItemTag = Empty
        End With
    Case TypeName(udtInfo)
        With varItem
            .BackgroundColor = Empty
            .BalloonLegendType = Empty
            .BalloonValue = Empty
            .Font = Empty
            .HotspotFont = Empty
            .HotspotColor = Empty
            .LegendType = Empty
            .NumberValue = Empty
            .RowIndex = Empty
            .RowNumber = Empty
            .ShadowLegendType = Empty
            .ShadowValue = Empty
            .Tag = Empty
            .TextColor = Empty
            .Time = Empty
            .TimeEnd = Empty
            .Value = Empty
            .Tip = Empty
            .Group = Empty
            .ShowHotspotCursor = False
        End With
    End Select
End Sub

Private Sub mmessageManager_ErrorShow(ByVal sender As Variant, ByVal e As ZLSoft_BusinessHome_ClientControl_TimeLineBase.IEventArgsMessageInfo)
    Dim objE As Object          '���Win Server 2003\XP ��δ��װ.NET�������ɲ�����ʾ����DLL����
    Set objE = e
    LogWrite "סԺһ���ĵ�����־", "" & glngModul, "mmessageManager_ErrorShow", "mmessageManager_ErrorShow:" & objE.Caption & vbCrLf & objE.Message & vbCrLf & objE.Exception
End Sub

Private Sub mmessageManager_InfoShow(ByVal sender As Variant, ByVal e As ZLSoft_BusinessHome_ClientControl_TimeLineBase.IEventArgsMessageInfo)
    Dim objE As Object          '���Win Server 2003\XP ��δ��װ.NET�������ɲ�����ʾ����DLL����
    Set objE = e
    LogWrite "סԺһ���ĵ�����־", "" & glngModul, "mmessageManager_InfoShow", "mmessageManager_InfoShow:" & objE.Caption & vbCrLf & objE.Message & vbCrLf & objE.Exception
End Sub

Private Sub mtimeLineControl_DataMouseClick(ByVal sender As Variant, ByVal e As ZLSoft_BusinessHome_ClientControl_TimeLineBase.IEventArgsDataInfo)
    Dim objPopup As CommandBarPopup
    Dim strTmp As String, strType As String
    Dim strMsg As String, str��鱨��ID As String
    Dim lng����ID As Long, lngҽ��ID As Long
    Dim objE As Object          '���Win Server 2003\XP ��δ��װ.NET�������ɲ�����ʾ����DLL����
    Set objE = e
    If objE.MouseButtons = MouseButtons_Left Then
        Select Case objE.Name
        Case "��������"
            strTmp = objE.Tag
            If strTmp = "" Then Exit Sub
            If Len(strTmp) < 32 Then
                lng����ID = CLng(strTmp) '�ϰ没���鿴
                Call gobjRichEPR.ViewDocument(Me, lng����ID, False)
            ElseIf Len(strTmp) = 32 And Not gobjEmr Is Nothing Then
                '�°没��
                On Error Resume Next
                strMsg = gobjEmr.OpenInEPR(strTmp)
                err.Clear: On Error GoTo 0
            End If
         Case "���", "����"
            '���ı���
            strTmp = objE.Tag
            strType = Split(strTmp, ",")(0)
            lng����ID = Val(Split(strTmp, ",")(1))
            lngҽ��ID = Val(objE.RowIndex)
            If objE.Name = "���" Then
                str��鱨��ID = Split(strTmp, ",")(2)
                Call FuncEPRReport(Me, lngҽ��ID, "D", lng����ID, str��鱨��ID, 2)
            ElseIf objE.Name = "����" Then
                Call FuncEPRReport(Me, lngҽ��ID, "", lng����ID, , 2)
            End If
        End Select
    End If
    
    If objE.MouseButtons = MouseButtons_Right Then
        If objE.Name = "���" Then
            mlngҽ��ID = Val(objE.RowIndex)
            Set objPopup = cbsSub.ActiveMenuBar.FindControl(, conMenu_EditPopup)
            If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Public Sub FuncCreateTimeLine()
'���ܣ���������
'����:bytFunc =1   ȱʡ��ʽ
'     bytFunc =2   ���ݸ���
    Dim strTitle As String
    Dim strDesignInfo As String
    Dim strBegin As String, strEnd As String
    Dim strSQL As String, strKey As String
    Dim i As Long, j As Long, lngPosition As Long
    Dim rsNurse As ADODB.Recordset       '���¼�¼��Ŀ
    Dim rsTemp As ADODB.Recordset
    Dim udtDataInfo As DataInfo
    Dim udtDataItem As DataItem
    
    Dim colTmp As Collection, colHead As Collection
    Dim colMeaSure As Collection
    Dim colFoot As Collection
    Dim colItem As Collection
    
    Dim strFile As String
    Dim strData As String
    Dim datBegin As Date, datEnd As Date
    
    On Error GoTo errH
    strEnd = Format(mDatEnd, "YYYY-MM-DDTHH:MM:SS")
    strBegin = Format(mDatBegin, "YYYY-MM-DDTHH:MM:SS")
    
    datBegin = CDate(Format(mDatBegin, "YYYY-MM-DD 00:00:00"))
    datEnd = CDate(Format(mDatEnd, "YYYY-MM-DD 23:59:59"))
    '��ʼ��ʱ�������
    '���¼�¼��Ŀ.��¼Ƶ��  ΪNULLʱĬ��Ϊ2
    strSQL = "Select a.��Ŀ���, a.�������, a.��¼��, a.��¼��, a.��¼��, a.��¼ɫ, a.���ֵ, a.��Сֵ, a.��λֵ, Nvl(a.��λ, b.��Ŀ��λ) As ��λ," & vbNewLine & _
            " NVL(a.��¼Ƶ��,2) as ��¼Ƶ��, a.�̶ȼ��, b.��Ŀ����, b.Ӧ�÷�ʽ, Decode(c.��Ŀ���, Null, Decode(d.���, Null, 0, 2), 1) As ��Ŀ����, Nvl(d.�����, 0) As ����� " & vbNewLine & _
            "From ���¼�¼��Ŀ A, �����¼��Ŀ B, ��������Ŀ C, ���������Ŀ D " & vbNewLine & _
            "Where a.��Ŀ��� = b.��Ŀ��� And b.��Ŀ��� = c.��Ŀ���(+) And b.��Ŀ��� = d.���(+) And (b.��Ŀ���� = 1 Or b.��Ŀ���� = 2 And Exists" & vbNewLine & _
            "       (Select 1" & vbNewLine & _
            "       From ���˻����ļ� A, �����ļ��б� E, ���˻������� C, ���˻�����ϸ D" & vbNewLine & _
            "       Where a.��ʽid = e.Id And a.Id = c.�ļ�id And c.Id = d.��¼id And a.����id = [1] And a.��ҳid = [2] And" & vbNewLine & _
            "       Nvl(a.Ӥ��, 0) = [3] And d.��Ŀ��� = b.��Ŀ��� And e.���� = 3 And e.���� = -1 And" & vbNewLine & _
            "       c.����ʱ�� Between [4] And [5])) And b.Ӧ�÷�ʽ > 0 And (b.���ÿ��� = 1 Or (b.���ÿ��� = 2 And Exists (Select 1 From �������ÿ��� Where ����id = [6]))) And Instr([7], ',' || b.���ò��� || ',') > 0" & vbNewLine & _
            "Order By ��¼��, �������, ��Ŀ���"
            
    Set rsNurse = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, mintBaby, datBegin, datEnd, mlng����ID, IIf(mintBaby > 0, ",0,2,", ",0,1,"))
    
    With mudtDesign
        .BackgroundColor = "200,255,255,255"
        .DateTitle = "����"
        .DateTitleFont = "����,9pt"
        .DateStart = strBegin
        .dateEnd = strEnd
        .ShowTick = IIf(mblnMeasureArea, True, False)
        .ShowFullDate = True
        .TickTitle = "ʱ��"
        .TickTitleFont = "����,9pt"
        .TickWidth = 20
        .TickRangeListCount = TICK_6
        .DateFont = "����,9pt"
        .TickFont = "����,9pt"
        .MergePeriodWidth = 50
        .EmptyDataMergeDayCount = 10    '����9������������ϲ�
        .EmptyDataMergePeriodWidth = 50
        .PaddingLeft = 10
        .PaddingTop = 10
        .PaddingRight = 10
        .PaddingBottom = 10
        .RowHeight = 22
        .MeasureTitleWidth = 60
        .GridMinHeight = 220
        .TopFixedSmallRowCount = 1
        .BottomFixedSmallRowCount = 1
        .GridYSplitCount = 5
        .GridYSmallSplitCount = 1
        .Height = 400
    End With
    'סԺ����---------------------------------------------
    Set colHead = New Collection
    FuncClearUDT udtDataItem
    With udtDataItem
        .GraphType = M_CON_TYPE_NORMAL
        .Title = "סԺ����"
        .TitleFont = "����,9pt"
        .TextColor = "Black"
        .Font = "����,9pt"
        Set .ListData = colTmp
    End With
    colHead.Add udtDataItem, M_CON_KEY_סԺ����
    '���������� ��������---------------------------------------------
    Call FuncClearUDT(udtDataItem)
    With udtDataItem
        .GraphType = M_CON_TYPE_COUNT
        .Title = "����������"
        .TitleFont = "����,9pt"
        .TextColor = "Black"
        .Font = "����,9pt"
        Set .ListData = colTmp
    End With
    colHead.Add udtDataItem, M_CON_KEY_����������
    'MeasureData ���---------------------------------------------
    Set colMeaSure = New Collection
    '���� ���ݻ����¼��Ŀ��������������Ƿ���ʾ����
    rsNurse.Filter = "��Ŀ���=-1"
    If Not rsNurse.EOF Then mlngӦ�÷�ʽ = Nvl(rsNurse!Ӧ�÷�ʽ, 0)

    rsNurse.Filter = "��¼��=1 And Ӧ�÷�ʽ=1" '������� ����\����\����
    For i = 1 To rsNurse.RecordCount
        FuncClearUDT udtDataItem
        With udtDataItem
            .GraphType = M_CON_TYPE_MEASURE
            .Title = rsNurse!��¼��
            .TitleFont = "����,9pt"
            .TextColor = "Black"
            .Unit = rsNurse!��λ & ""
            .MinValue = rsNurse!��Сֵ
            .MaxValue = rsNurse!���ֵ
            j = (rsNurse!���ֵ - rsNurse!��Сֵ) / Nvl(rsNurse!�̶ȼ��, 1)
            If rsNurse!��¼�� = "����" Then
                j = 5
            ElseIf rsNurse!��¼�� = "����" Or rsNurse!��¼�� = "����" Then
                j = 7
            Else
                j = (rsNurse!���ֵ - rsNurse!��Сֵ) / Nvl(rsNurse!�̶ȼ��, 1)
            End If
            .SplitNum = j
            .Color = FuncColorRGB(CLng(rsNurse!��¼ɫ & ""))
            .LegendType = "ʵ��Բ"  '��ʵ��Բ������Բ���ֿ���Բ���㡢�桢H���ţ�
            .ShadowLegendType = "����Բ"     '��Ӱ��
            .BalloonLegendType = "����Բ"    '�����
            .IsDataDynamicExpansion = True
            If rsNurse!��Ŀ��� = 1 Then
            '����
                .BalloonTitle = "����������"
                .BalloonColor = "Red"
            ElseIf rsNurse!��Ŀ��� = 2 Then
            '����
                .ShadowTitle = "����"
                If mlngӦ�÷�ʽ = 1 Or mlngӦ�÷�ʽ = 2 Then
                    lngPosition = rsNurse.AbsolutePosition
                    '���ʵ���Ӧ��ʱ;��������Ӱ������ɫȡ���ʵ���ɫֵ
                    rsNurse.Filter = "��Ŀ���=-1"
                    If Not rsNurse.EOF Then
                        .ShadowColor = FuncColorRGB(CLng(rsNurse!��¼ɫ & ""))
                    End If
                    '�ָ���ԭ��ָ��λ��
                    rsNurse.Filter = "��¼��=1 And Ӧ�÷�ʽ=1" '������� ����\����\����
                    rsNurse.AbsolutePosition = lngPosition
                End If
            End If
            .ItemTag = "K_" & rsNurse!��Ŀ���
        End With
        colMeaSure.Add udtDataItem, "K_" & rsNurse!��Ŀ���
        rsNurse.MoveNext
    Next
    '-------------------------------------------------------------------------------------------------------
    Set colFoot = New Collection
    
    Set colItem = New Collection
    rsNurse.Filter = "��¼��=2 And Ӧ�÷�ʽ=1 And ��Ŀ��� <> 4 " '3-����\Ѫѹ(4-����ѹ/5-����ѹ)\7-��Һ��\9-��Һ��\10-������\�Զ�����Ŀ..
    For i = 1 To rsNurse.RecordCount
        FuncClearUDT udtDataItem
        strKey = "K_" & rsNurse!��Ŀ���
        strTitle = rsNurse!��¼�� & IIf(Nvl(rsNurse!��λ) <> "", "(" & rsNurse!��λ & ")", "")
        With udtDataItem
            If rsNurse!��Ŀ��� = 3 Then
                .GraphType = M_CON_TYPE_TICK    'ʱ������
            ElseIf Val(rsNurse!��¼Ƶ�� & "") > 0 Then
                .GraphType = M_CON_TYPE_DIVISION
                .SplitString = ","
                .SplitCount = Val(rsNurse!��¼Ƶ�� & "")
                If rsNurse!��Ŀ��� = 5 Then
                    strTitle = "Ѫѹ" & IIf(Nvl(rsNurse!��λ) <> "", "(" & rsNurse!��λ & ")", "")
                End If
            End If
            .Title = strTitle
            .TitleFont = "����,9pt"
            .TextColor = IIf(Val(rsNurse!��¼ɫ & "") = 0, "Black", FuncColorRGB(Val(rsNurse!��¼ɫ & "")))
            .BackgroundColor = "255,255,255"
            .Font = "����,9pt"
            .ItemTag = strKey & "," & rsNurse!��Ŀ��� & "," & rsNurse!��Ŀ���� & "," & rsNurse!�����    '����,��Ŀ���,��Ŀ����(0-�ձ�,1-����,2-����),�����
        End With
        colItem.Add udtDataItem, strKey
        rsNurse.MoveNext
    Next
    
    FuncClearUDT udtDataItem
    With udtDataItem
        .GraphType = M_CON_TYPE_GROUP
        .Title = "������Ŀ"
        .GroupPosition = "����"    '����\����
        Set .ListData = colItem
    End With
    colFoot.Add udtDataItem, "K_������Ŀ"
    '--------------------------------------------------------------------------------------------------------------
    '��ҩ����
    FuncClearUDT udtDataItem
    With udtDataItem
        .GraphType = M_CON_TYPE_CONTINUOUS
        .Title = "ҩƷ����"
        .TitleFont = "����,9pt"
        .TextColor = "Black"
        .Font = "����,9pt"
        .Effect = "����"
        .ShowHotspotEffect = True
        .HotspotColor = FuncColorRGB(COLOR_����)
        .HotspotFont = "����,9pt"
    End With
    colFoot.Add udtDataItem, M_CON_KEY_��ҩ����_��
    
    FuncClearUDT udtDataItem
    With udtDataItem
        .GraphType = M_CON_TYPE_CONTINUOUS
        .Title = "ҩƷ����"
        .TitleFont = "����,9pt"
        .TextColor = "Black"
        .Font = "����,9pt"
        .Effect = "����"
        .ShowHotspotEffect = True
        .HotspotColor = FuncColorRGB(COLOR_����)
        .HotspotFont = "����,9pt"
    End With
    colFoot.Add udtDataItem, M_CON_KEY_��ҩ����_��
    '���
    FuncClearUDT udtDataItem
    With udtDataItem
        .GraphType = M_CON_TYPE_LAYOUT
        .Title = "���"
        .TitleFont = "����,9pt"
        .TextColor = "Black"
        .BackgroundColor = "255,255,255"
        .Font = "����,9pt"
        .ShowHotspotEffect = True
        .HotspotFont = "����,9pt,style=Underline"
        .HotspotColor = FuncColorRGB(COLOR_����)
    End With
    colFoot.Add udtDataItem, M_CON_KEY_���
    '����
    FuncClearUDT udtDataItem
    With udtDataItem
        .GraphType = M_CON_TYPE_LAYOUT
        .Title = "����"
        .TitleFont = "����,9pt"
        .TextColor = "Black"
        .BackgroundColor = "255,255,255"
        .Font = "����,9pt"
        .ShowHotspotEffect = True
        .HotspotFont = "����,9pt,style=Underline"
        .HotspotColor = FuncColorRGB(COLOR_����)
    End With
    colFoot.Add udtDataItem, M_CON_KEY_����
    '����ҽ��:��ʱ������
    FuncClearUDT udtDataItem
    With udtDataItem
        .GraphType = M_CON_TYPE_LAYOUT
        .Title = "��������"
        .TitleFont = "����,9pt"
        .TextColor = "Black"
        .BackgroundColor = "White"
        .Font = "����,9pt"
        .ShowHotspotEffect = True
        .HotspotColor = FuncColorRGB(COLOR_����)
        .HotspotFont = "����,9pt"
    End With
    colFoot.Add udtDataItem, M_CON_KEY_��������
    
    FuncClearUDT udtDataItem
    With udtDataItem
        .GraphType = M_CON_TYPE_CONTINUOUS
        .Title = "��������"
        .TitleFont = "����,9pt"
        .TextColor = "Black"
        .BackgroundColor = "White"
        .Font = "����,9pt"
        .ShowHotspotEffect = True
        .HotspotColor = FuncColorRGB(COLOR_����)
        .HotspotFont = "����,9pt"
        .Effect = "����"
    End With
    colFoot.Add udtDataItem, M_CON_KEY_��������
    
    '����
    FuncClearUDT udtDataItem
    With udtDataItem
        .GraphType = M_CON_TYPE_LAYOUT
        .Title = "����"
        .TitleFont = "����,9pt"
        .TextColor = "Black"
        .BackgroundColor = "White"
        .Font = "����,9pt"
        .ShowHotspotEffect = True
        .HotspotColor = FuncColorRGB(COLOR_����)
        .HotspotFont = "����,9pt"
    End With
    colFoot.Add udtDataItem, M_CON_KEY_����
    
    '��������
    FuncClearUDT udtDataItem
    With udtDataItem
        .GraphType = M_CON_TYPE_LAYOUT
        .Title = "��������"
        .TitleFont = "����,9pt"
        .TextColor = "Black"
        .BackgroundColor = "White"
        .Font = "����,9pt"
        .ShowHotspotEffect = True
        .HotspotFont = "����,9pt,style=Underline"
        .HotspotColor = FuncColorRGB(COLOR_���)
        .ShowHotspotCursor = True
        .GroupPosition = "����"
    End With
    colFoot.Add udtDataItem, M_CON_KEY_����
    
    '---------------------------------------------
    With mudtTimeLine
        Set .colHeaderData = colHead
        Set .colMeasureData = colMeaSure
        Set .colFooterData = colFoot
    End With
    '----------------------------------------------------------------------------------------------------------------------------------------

    If mlng����ID <> 0 Then
        Call FuncLoadItemNurse
        Call FuncLoadDrug
        Call FuncLoadEMR
        Call FuncLoadPACS_Lis
        Call FuncLoadOperation
        Call FuncLoadInHosDay
        Call FuncLoadOperationAfter
        Call FuncLoadAdvice
    End If
    
    strDesignInfo = FuncMakeXMLDesign(mudtDesign)
    LogWrite "סԺһ���ĵ�����־", "" & glngModul, "FuncCreateTimeLine", "Design_Create:" & vbCrLf & strDesignInfo
    strData = FuncMakeXMLTimeLine(mudtTimeLine)
    LogWrite "סԺһ���ĵ�����־", "" & glngModul, "FuncCreateTimeLine", "Data_Create:" & vbCrLf & strData
    mtimeLineControl.UpdateDesignInfo strDesignInfo
    mtimeLineControl.UpdateData strData
    mtimeLineControl.RefreshAll
    If mblnMeasureArea Then
        mtimeLineControl.ShowMeasureArea
    Else
        mtimeLineControl.HideMeasureArea
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitCommandBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objMenu As CommandBarPopup
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsSub.VisualTheme = xtpThemeOfficeXP
    With Me.cbsSub.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .UseFadedIcons = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With

    Set cbsSub.Icons = zlCommFun.GetPubIcons
    cbsSub.EnableCustomization False
    cbsSub.ActiveMenuBar.Visible = False
    '------------------------------------------------------------------------------------------------------------------
    
    '�����˵�
    Set objMenu = cbsSub.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�����˵�(&K)", 0, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Img_Look, "��Ƭ")
        objControl.IconId = conMenu_Img_Look
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '����������:������������
    Set objBar = cbsSub.Add("������", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap '+ xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False

    With objBar.Controls
    
        Set mobjPopup = .Add(xtpControlPopup, conMenu_Edit_NewItem, "ҳ��")
        mobjPopup.IconId = conMenu_Edit_Modify
        mobjPopup.Style = xtpButtonIconAndCaption
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Forward, "��һҳ", -1, False)
        objControl.IconId = conMenu_View_Forward
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Backward, "��һҳ", -1, False)
        objControl.IconId = conMenu_View_Backward
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Show, IIf(mblnMeasureArea, "���ػ���", "��ʾ����"), -1, False)
        objControl.IconId = conMenu_Manage_Up
        objControl.Style = xtpButtonIconAndCaption
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Assistant, "ʮ����", -1, False)
        objControl.IconId = conMenu_PatholMeal_AddRecord
        objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, conMenu_Process_Zoom, "�Ŵ�", -1, False)
        objControl.IconId = conMenu_Process_Zoom
        objControl.Style = xtpButtonIconAndCaption
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Navigatebeginning, "", -1, False)
        objControl.Style = xtpButtonIconAndCaption
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Navigateend, "", -1, False)
        objControl.Style = xtpButtonIconAndCaption
    End With
End Sub

Private Sub FuncLoadPages()
'����:����ʵ�����,���ҳ�水ť
    Dim objControl As CommandBarControl
    Dim datDate As Date
    Dim strSQL As String, strTmpDate As String
    Dim i As Long, lngDay As Long
    Dim rsTmp As ADODB.Recordset
    
    If Not mobjPopup Is Nothing Then mobjPopup.CommandBar.Controls.DeleteAll
    
    datDate = zlDatabase.Currentdate
    If mlng����ID = 0 Then
        'δѡ�в���ʱ , ����ȱʡ����
        Set objControl = mobjPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_View_Jump, "��1ҳ:" & Format(datDate, "YYYY-MM-DD") & "��" & Format(datDate + 6, "YYYY-MM-DD"), -1, False)
        objControl.Parameter = 1 & "," & Format(datDate, "YYYY-MM-DD") & "," & Format(datDate + 6, "YYYY-MM-DD")
        mlngPages = 1
    Else
        '��ȡ������Ժ����,��Ժ����
        strSQL = "Select a.��Ժ����, a.��Ժ���� From ������ҳ A Where a.����id = [1] And a.��ҳid = [2]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
        mDatIn = CDate(Format(rsTmp!��Ժ����, "YYYY-MM-DD"))
        mdatOut = CDate(Format(IIf(Nvl(rsTmp!��Ժ����, 0) = 0, Format(datDate, "YYYY-MM-DD"), Format(rsTmp!��Ժ����, "YYYY-MM-DD")), "YYYY-MM-DD"))
        lngDay = (mdatOut - mDatIn) + 1
        If lngDay <= 0 Then
            mlngPages = 1
            mlngDay = 7
        Else
            mlngDay = 7
            mlngPages = lngDay \ mlngDay + IIf(lngDay Mod mlngDay > 0, 1, 0)
        End If
        For i = 1 To mlngPages
            mDatBegin = mDatIn + (i - 1) * mlngDay
            If i < mlngPages Then
                mDatEnd = mDatBegin + (mlngDay - 1)
            Else
                lngDay = mdatOut - mDatBegin
                If lngDay < 7 Then
                    mDatEnd = mDatBegin + 6  '
                Else
                    mDatEnd = mdatOut
                End If
            End If
            strTmpDate = Format(mDatBegin, "YYYY-MM-DD") & "��" & Format(mDatEnd, "YYYY-MM-DD")
            Set objControl = mobjPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_View_Jump, "��" & i & "ҳ:" & strTmpDate, -1, False)
            objControl.Parameter = i & "," & Format(mDatBegin, "YYYY-MM-DD") & "," & Format(mDatEnd, "YYYY-MM-DD")
        Next
    End If
    objControl.Execute
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncLoadDrug()
'����:
    Dim strSQL As String, strType As String, strColor As String
    Dim str��Ч As String, strBegin As String, strEnd As String
    Dim strValue As String, strTag As String, str�÷� As String
    Dim strName As String
    Dim rsAdvice As ADODB.Recordset, rsDrug As ADODB.Recordset
    Dim colFoot As Collection
    Dim udtDataInfo As DataInfo
    Dim udtDataItem As DataItem
    Dim i As Long, j As Long, n As Long, k As Long, lngPos As Long, lng��ID As Long
    Dim lngRow As Long, lngDay As Long, lngTemp As Long, lngGroupNum As Long
    Dim datBegin As Date, datEnd As Date, datCurr As Date, datTemp As Date
    Dim blnGroup As Boolean
    
    On Error GoTo errH
        
    datBegin = CDate(Format(mDatBegin, "YYYY-MM-DD 00:00:00"))
    datEnd = CDate(Format(mDatEnd, "YYYY-MM-DD 23:59:59"))
    datCurr = zlDatabase.Currentdate
    '��ҩ����ȡ1-��ҩ���г�ҩ����ҽ����¼��Ƥ��ҽ����2-��ҩ�䷽ֻȡ��ҩ���� �������=E,�������� =4
    strSQL = "Select a.Id, a.���id, a.ҽ����Ч, a.�������, b.����, a.ҽ������, b.ִ�з���, b.��������, a.����ʱ��, a.��ʼִ��ʱ��, a.ִ����ֹʱ��," & vbNewLine & _
            "       Decode(a.�״�����, Null, '', a.�״����� || b.���㵥λ || ':') ||" & vbNewLine & _
            "        Decode(a.��������, Null, Null, Decode(Sign(1 - a.��������), 1, '0' || a.��������, a.��������) || b.���㵥λ) As ����," & vbNewLine & _
            "       Decode(a.�ܸ�����, Null, Null," & vbNewLine & _
            "               Decode(a.�������, 'E', Decode(b.��������, '4', a.�ܸ����� || '��', a.�ܸ����� || b.���㵥λ), '5'," & vbNewLine & _
            "                       Round(a.�ܸ����� / d.סԺ��װ, 5) || d.סԺ��λ, '6', Round(a.�ܸ����� / d.סԺ��װ, 5) || d.סԺ��λ, a.�ܸ����� || b.���㵥λ)) As ����," & vbNewLine & _
            "       a.ִ��Ƶ��, Decode(A.�������,'E',Decode(Instr('2468',Nvl(B.��������,'0')),0,NULL,B.����),NULL) as �÷� " & vbNewLine & _
            "From ����ҽ����¼ A, ������ĿĿ¼ B, ҩƷ��� D" & vbNewLine & _
            "Where a.������Ŀid = b.Id(+) And a.�շ�ϸĿid = d.ҩƷid(+) And a.����id = [1] And a.��ҳid = [2] And Nvl(a.Ӥ��, 0) = [3] And" & vbNewLine & _
            "      (Instr(',5,6,', a.�������) > 0 Or (a.������� = 'E' And b.�������� In ('1', '2', '4'))) And" & vbNewLine & _
            "      (a.��ʼִ��ʱ�� Between [4] And [5] OR (a.ҽ����Ч=0 And a.��ʼִ��ʱ�� < [4] And NVL(a.ִ����ֹʱ��,[5])>[4]))  And" & vbNewLine & _
            "      ((a.ҽ����Ч = 1 And a.ҽ��״̬ In (3, 8)) Or (a.ҽ����Ч = 0 And a.ҽ��״̬ In (3, 5, 6, 7, 8, 9)))" & vbNewLine & _
            "Order By a.ҽ����Ч, a.��ʼִ��ʱ��, a.���"

    Set rsAdvice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, mintBaby, datBegin, datEnd)
    For n = 1 To 2
        If n = 1 Then
            rsAdvice.Filter = "ҽ����Ч=0"
            udtDataItem = mudtTimeLine.colFooterData(M_CON_KEY_��ҩ����_��)
        ElseIf n = 2 Then
            rsAdvice.Filter = "ҽ����Ч=1"
            udtDataItem = mudtTimeLine.colFooterData(M_CON_KEY_��ҩ����_��)
        End If
        Set colFoot = New Collection
        lngRow = 0: Set rsDrug = InitRS()
        For i = 1 To rsAdvice.RecordCount
            lngPos = rsAdvice.AbsolutePosition '��ǵ�ǰ��
            blnGroup = False
            If (rsAdvice!�������� & "" = "4" And rsAdvice!������� & "" = "E") Or InStr(",5,6,", rsAdvice!�������) > 0 Then
                If InStr(",5,6,", rsAdvice!�������) > 0 Then
                    If lng��ID <> CLng(rsAdvice!���ID & "") Then
                        blnGroup = True
                        lng��ID = CLng(rsAdvice!���ID & "")
                        rsAdvice.MoveNext
                        For j = lngPos + 1 To rsAdvice.RecordCount
                            If lng��ID = CLng(rsAdvice!ID & "") Then
                                '0-�����������,1-��Һ��,2-ע����,3-Ƥ��,4-�ڷ�
                                Select Case Nvl(rsAdvice!ִ�з���)
                                Case "1"
                                    strType = "[��]"      'FFC0C0 =BGR   ͬ����һ������һ��
                                    strColor = FuncColorRGB(COLOR_ǳ��)
                                Case "2"
                                    strType = "[��]"       '&HFFFFC0
                                    strColor = FuncColorRGB(COLOR_ǳ��)
                                Case "3"
                                    strType = "[Ƥ]"      '&HC0C0FF
                                    strColor = FuncColorRGB(COLOR_ǳ��)
                                Case "4"
                                    strType = "[��]"        '&HC0FFC0
                                    strColor = FuncColorRGB(COLOR_ǳ��)
                                Case Else
                                    strType = ""
                                    strColor = FuncColorRGB(COLOR_ǳ��)           '&HC0E0FF
                                End Select
                                str�÷� = rsAdvice!�÷� & ""
                                lngGroupNum = j - lngPos
                                Exit For
                            End If
                            rsAdvice.MoveNext
                        Next
                    End If
                    rsAdvice.AbsolutePosition = lngPos
                    strTag = "ҩ������:" & rsAdvice!ҽ������ & IIf(Nvl(rsAdvice!����) = "", "", ",��" & rsAdvice!����) & IIf(Nvl(rsAdvice!����) = "", "", ",ÿ��" & rsAdvice!����) & "," & str�÷� & "," & rsAdvice!ִ��Ƶ��
                            
                    strName = rsAdvice!���� & ""
                ElseIf (rsAdvice!�������� & "" = "4" And rsAdvice!������� & "" = "E") Then
                    strType = "��"
                    strTag = "ҩ������:" & rsAdvice!ҽ������
                    strName = rsAdvice!ҽ������ & ""
                    strColor = FuncColorRGB(COLOR_ǳ��)
                End If
                
                strTag = strTag & vbCrLf & "��Чʱ��:" & Format(rsAdvice!��ʼִ��ʱ�� & "", "YYYY-MM-DD HH:MM:SS")
                
                If Val(rsAdvice!ҽ����Ч & "") = 1 Then
                    strEnd = ""
                    strValue = "[��]" & strType & strName
                    strColor = ""
                    strBegin = Format(rsAdvice!��ʼִ��ʱ��, "YYYY-MM-DD ") & Format("2:30:00", "HH:MM:SS")
                    '������ʱҩ����ʾ��ռ����,�������һ��ʱ,Ĭ����ֹʱ��Ϊ���һ�е�ʱ��
                    If rsAdvice!ִ����ֹʱ�� & "" = "" Then
                        datTemp = DateAdd("D", 1, CDate(rsAdvice!��ʼִ��ʱ�� & ""))
                    Else
                        datTemp = CDate(rsAdvice!ִ����ֹʱ�� & "")
                        lngDay = DateDiff("D", CDate(rsAdvice!��ʼִ��ʱ�� & ""), CDate(rsAdvice!ִ����ֹʱ�� & ""))
                        If lngDay < 1 Then
                            datTemp = CDate(rsAdvice!ִ����ֹʱ�� & "") + 1
                        End If
                    End If
                    If datTemp > datEnd Then
                        strEnd = Format(datEnd, "YYYY-MM-DD 23:59:59")
                    Else
                        strEnd = Format(datTemp, "YYYY-MM-DD 23:59:59")
                    End If
                Else
                    If rsAdvice!ִ����ֹʱ�� & "" <> "" Then
                        strTag = strTag & vbCrLf & "��ֹʱ��:" & Format(rsAdvice!ִ����ֹʱ�� & "", "YYYY-MM-DD HH:MM:SS")
                    End If
                    If Between(CDate(Format(rsAdvice!��ʼִ��ʱ�� & "", "YYYY-MM-DD")), datBegin, datEnd) Then
                        strBegin = Format(rsAdvice!��ʼִ��ʱ��, "YYYY-MM-DD ") & Format("2:30:00", "HH:MM:SS")
                    Else
                        strBegin = Format(datBegin, "YYYY-MM-DD ") & Format("2:30:00", "HH:MM:SS")
                    End If
                    
                    If Nvl(rsAdvice!ִ����ֹʱ��) = "" Then
                        If DateDiff("D", datEnd, datCurr) >= 0 Then
                            datTemp = datEnd
                        Else
                            datTemp = datCurr
                        End If
                    Else
                        If CDate(rsAdvice!ִ����ֹʱ��) > datEnd Then
                            datTemp = datEnd
                        Else
                            datTemp = CDate(rsAdvice!ִ����ֹʱ��)
                        End If
                    End If
                    strEnd = Format(datTemp, "YYYY-MM-DD 23:59:59")
                    strValue = "[��]" & strType & rsAdvice!����
                End If
                
                '��ҩ���ٿհ��
                If blnGroup Then
                    lngTemp = 0
                    For j = 1 To lngRow
                        rsDrug.Filter = "�к�=" & j & " And ���� >= '" & Format(strBegin, "YYYY-MM-DD") & "'"
                        If rsDrug.RecordCount = 0 Then
                            lngTemp = j   '�ҵ�һ�пհ���
                            If lngGroupNum <= 1 Then
                                Exit For
                            Else
                                'һ����ҩ�ж�Ԥ���հ��Ƿ��㹻
                                For k = 2 To lngGroupNum
                                    rsDrug.Filter = "�к�=" & (j + k - 1) & " And ���� >= '" & Format(strBegin, "YYYY-MM-DD") & "'"
                                    If rsDrug.RecordCount > 0 Then Exit For
                                Next
                                If k <= lngGroupNum Then
                                    j = lngRow + 1: lngTemp = lngRow
                                End If
                                Exit For
                            End If
                        End If
                    Next
                Else
                    If lngTemp < lngRow And lngTemp > 0 Then lngTemp = lngTemp + 1
                End If
                If lngTemp = lngRow Or lngTemp = 0 Then lngRow = lngRow + 1: lngTemp = lngRow
                 
                If Val(rsAdvice!ҽ����Ч & "") = 1 Then
                    lngDay = 1
                Else
                    lngDay = DateDiff("D", Format(strBegin, "YYYY-MM-DD"), Format(strEnd, "YYYY-MM-DD"))
                    If lngDay < 1 Then lngDay = 1
                End If
                rsDrug.AddNew
                For j = 0 To lngDay
                    rsDrug!�к� = lngTemp
                    rsDrug!���� = Format(DateAdd("D", j, Format(strBegin, "YYYY-MM-DD")), "YYYY-MM-DD")
                Next
                rsDrug.UpdateBatch
                
                With udtDataInfo
                    .Value = strValue
                    .RowNumber = lngTemp
                    .RowIndex = rsAdvice!ID
                    .Time = Format(strBegin, "YYYY-MM-DDTHH:MM:SS")
                    .TimeEnd = Format(strEnd, "YYYY-MM-DDTHH:MM:SS")
                    .BackgroundColor = strColor
                    .Group = rsAdvice!���ID & "" '������Ϣ��¼���ID
                    .Tip = strTag
                End With
                colFoot.Add udtDataInfo, "_" & (colFoot.Count + 1)
            End If
            rsAdvice.MoveNext
        Next
        If n = 1 Then
            Set udtDataItem.ListData = colFoot
            mudtTimeLine.colFooterData.Remove (M_CON_KEY_��ҩ����_��)
            mudtTimeLine.colFooterData.Add udtDataItem, M_CON_KEY_��ҩ����_��, M_CON_KEY_���
        Else
            Set udtDataItem.ListData = colFoot
            mudtTimeLine.colFooterData.Remove (M_CON_KEY_��ҩ����_��)
            mudtTimeLine.colFooterData.Add udtDataItem, M_CON_KEY_��ҩ����_��, M_CON_KEY_��ҩ����_��
        End If
    Next

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncLoadInHosDay()
'����:����סԺ���� ��ͨ�ı�
'�㷨:����Ժ���ڵ���ǰ����
    Dim colHead As Collection
    Dim udtDataInfo As DataInfo
    Dim udtDataItem As DataItem
    Dim lngDay As Long
    Dim i As Long
    
    udtDataItem = mudtTimeLine.colHeaderData(M_CON_KEY_סԺ����)
    Set colHead = New Collection
    lngDay = (mdatOut - mDatIn)
    For i = 0 To lngDay
        With udtDataInfo
            .Value = mDatBegin - mDatIn + i
            .Time = Format(mDatBegin + i, "YYYY-MM-DDTHH:MM:SS")
            .BackgroundColor = "White"
            .TextColor = "Black"
        End With
        colHead.Add udtDataInfo, "_" & (colHead.Count + 1)
        If Format(mdatOut, "YYYY-MM-DD") = Format(mDatBegin + i, "YYYY-MM-DD") Then Exit For
    Next
    Set udtDataItem.ListData = colHead
    mudtTimeLine.colHeaderData.Remove (M_CON_KEY_סԺ����)
    mudtTimeLine.colHeaderData.Add udtDataItem, M_CON_KEY_סԺ����, M_CON_KEY_����������
    
End Sub

Private Sub FuncLoadPACS_Lis()
'����:������
    Dim colFoot As Collection
    Dim udtDataInfo As DataInfo
    Dim udtDataItem As DataItem
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim strType As String
    Dim strSQL As String
    Dim datBegin As Date, datEnd As Date
    
    datBegin = CDate(Format(mDatBegin, "YYYY-MM-DD 00:00:00"))
    datEnd = CDate(Format(mDatEnd, "YYYY-MM-DD 23:59:59"))
    
    strSQL = "Select a.Id, a.ҽ������, a.�������, b.��������, a.��ʼִ��ʱ��, Max(c.����id) As ����id, Max(c.��鱨��id) As ��鱨��id," & vbNewLine & _
            "       Decode(Max(Nvl(c.����״̬, 0)), Min(Nvl(c.����״̬, 0)), Max(Nvl(c.����״̬, 0)), 2) As ����״̬" & vbNewLine & _
            "From ����ҽ����¼ A, ������ĿĿ¼ B, ����ҽ������ C" & vbNewLine & _
            "Where a.������Ŀid = b.Id And a.Id = c.ҽ��id(+) And a.����id = [1] And a.��ҳid = [2] And NVL(a.Ӥ��,0) =[3] And" & vbNewLine & _
            "      ((a.ҽ����Ч = 1 And a.ҽ��״̬ In (3, 8)) Or (a.ҽ����Ч = 0 And a.ҽ��״̬ In (3, 5, 6, 7, 8, 9))) And" & vbNewLine & _
            "      (a.������� = 'D' And a.���id Is Null Or a.������� = 'E' And b.�������� = '6')" & vbNewLine & _
            "      And a.��ʼִ��ʱ�� Between [4] And [5] " & vbNewLine & _
            "Group By a.Id, a.ҽ������, a.���, a.�������, b.��������, a.��ʼִ��ʱ�� " & vbNewLine & _
            "Order By a.���"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, mintBaby, datBegin, datEnd)
    
    udtDataItem = mudtTimeLine.colFooterData(M_CON_KEY_���)
    Set colFoot = New Collection
    rsTmp.Filter = "������� = 'D'"
    For i = 1 To rsTmp.RecordCount
        Call FuncClearUDT(udtDataInfo)
        With udtDataInfo
            .RowIndex = rsTmp!ID
            .Value = rsTmp!ҽ������
            .Time = Format(rsTmp!��ʼִ��ʱ��, "YYYY-MM-DDTHH:MM:SS")
            .BackgroundColor = "White"
            '�ѳ����� ����ɫ������ʾ���ӱ���ͼ��
            If Not (Val(rsTmp!����ID & "") = 0 And Val(rsTmp!��鱨��ID & "") = 0) Then
                .TextColor = FuncColorRGB(COLOR_����)
                .ShowHotspotCursor = True
            Else
                .TextColor = "Black"
                .HotspotFont = "����,9pt"
                .ShowHotspotCursor = False
            End If
            .Tag = "D," & rsTmp!����ID & "," & rsTmp!��鱨��ID
        End With
        colFoot.Add udtDataInfo, "_" & (colFoot.Count + 1)
        rsTmp.MoveNext
    Next
    Set udtDataItem.ListData = colFoot
    mudtTimeLine.colFooterData.Remove (M_CON_KEY_���)
    mudtTimeLine.colFooterData.Add udtDataItem, M_CON_KEY_���, M_CON_KEY_����
    '����
    udtDataItem = mudtTimeLine.colFooterData(M_CON_KEY_����)
    Set colFoot = New Collection
    rsTmp.Filter = "������� = 'E'"
    For i = 1 To rsTmp.RecordCount
        With udtDataInfo
            .RowIndex = rsTmp!ID
            .Value = rsTmp!ҽ������
            .Time = Format(rsTmp!��ʼִ��ʱ��, "YYYY-MM-DDTHH:MM:SS")
            .BackgroundColor = "White"
            .Tag = "E," & rsTmp!����ID & "," & rsTmp!��鱨��ID
            '�ѳ����� ����ɫ������ʾ���ӱ���ͼ��
            If Val(rsTmp!����ID & "") <> 0 Then
                .TextColor = FuncColorRGB(COLOR_����)
                .ShowHotspotCursor = True
            Else
                .TextColor = "Black"
                .HotspotFont = "����,9pt"  'ȡ���»���
                .ShowHotspotCursor = False
            End If
        End With
        colFoot.Add udtDataInfo, "_" & (colFoot.Count + 1)
        rsTmp.MoveNext
    Next
    Set udtDataItem.ListData = colFoot
    mudtTimeLine.colFooterData.Remove (M_CON_KEY_����)
    mudtTimeLine.colFooterData.Add udtDataItem, M_CON_KEY_����, M_CON_KEY_��������
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
End Sub

Private Sub FuncLoadEMR()
'����:���ز����ļ��б�
    Dim colFoot As Collection
    Dim udtDataInfo As DataInfo
    Dim udtDataItem As DataItem
    Dim strSQL As String, strMsg As String
    Dim rsTmp As ADODB.Recordset
    Dim rsEmr As ADODB.Recordset
    Dim i As Long
    Dim datBegin As Date, datEnd As Date
    
    datBegin = CDate(Format(mDatBegin, "YYYY-MM-DD 00:00:00"))
    datEnd = CDate(Format(mDatEnd, "YYYY-MM-DD 23:59:59"))
     
    '2-סԺ����,5-�������,6-֪���ļ�
    strSQL = "Select ID, ��������, ��������,����ʱ��, Decode(Nvl(ǩ������, 0), 0, '����(δ���)', 1, '���', '��ǩ') as ״̬ " & vbNewLine & _
            "From ���Ӳ�����¼ " & vbNewLine & _
            "Where ������Դ = 2 And �������� In (2, 5, 6) And ����id = [1] And ��ҳid = [2] And NVL(Ӥ��,0) = [3] " & vbNewLine & _
            "      And ����ʱ�� Between [4] And [5] " & vbNewLine & _
            "Order By ��������, ���"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, mintBaby, datBegin, datEnd)
    
    udtDataItem = mudtTimeLine.colFooterData(M_CON_KEY_����)
    Set colFoot = New Collection
    
    For i = 1 To rsTmp.RecordCount
        With udtDataInfo
            .Value = rsTmp!�������� & "��" & rsTmp!״̬ & "��"
            .Time = Format(rsTmp!����ʱ�� & "", "YYYY-MM-DDTHH:MM:SS")
            .BackgroundColor = "White"
            .Tag = rsTmp!ID
        End With
        colFoot.Add udtDataInfo, "_" & (colFoot.Count + 1)
        rsTmp.MoveNext
    Next
    '�°没��
    If Not gobjEmr Is Nothing Then
        '�°没���ṩ�ӿڣ�GetInEPRRecord(����ID,��ҳID,RS)����ÿ�ξ���Ĳ��������Id, Title,Creat_Time,STATUS( �༭�С���ǩ�������С�����ǩ)����
        On Error Resume Next
        strMsg = gobjEmr.GetInEPRRecord(mlng����ID, mlng��ҳID, rsEmr)
        err.Clear: On Error GoTo 0
        If Not rsEmr Is Nothing Then
            For i = 1 To rsEmr.RecordCount
                With udtDataInfo
                    .Value = rsEmr!Title
                    .Time = Format(rsEmr!Creat_Time & "", "YYYY-MM-DDTHH:MM:SS")
                    If rsEmr.Fields.Count = 4 Then
                        If UCase(rsEmr.Fields(3).Name) = UCase("Status") Then
                            .Value = rsEmr!Title & IIf(rsEmr!Status & "" <> "", "��" & rsEmr!Status & "��", "")
                        End If
                    End If
                    .BackgroundColor = "White"
                    .Tag = rsEmr!ID   '�°没��IDֵ
                End With
                colFoot.Add udtDataInfo, "_" & (colFoot.Count + 1)
                rsEmr.MoveNext
            Next
        End If
    End If
    
    Set udtDataItem.ListData = colFoot
    mudtTimeLine.colFooterData.Remove (M_CON_KEY_����)
    mudtTimeLine.colFooterData.Add udtDataItem, M_CON_KEY_����
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
        
End Sub

Private Sub FuncLoadOperation()
'����:��������ҽ��
    Dim colFoot As Collection
    Dim udtDataInfo As DataInfo
    Dim udtDataItem As DataItem
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim datBegin As Date, datEnd As Date
    
    datBegin = CDate(Format(mDatBegin, "YYYY-MM-DD 00:00:00"))
    datEnd = CDate(Format(mDatEnd, "YYYY-MM-DD 23:59:59"))
    
    '������ҽ��
    strSQL = "Select a.Id, a.����ʱ��, a.ҽ������, b.���� " & vbNewLine & _
            "From ����ҽ����¼ A,������ĿĿ¼ B " & vbNewLine & _
            "Where a.������Ŀid = b.Id And a.����id = [1] And a.��ҳid = [2] And NVL(a.Ӥ��,0) =[3] And a.������� = 'F' And Nvl(a.���id, 0) = 0 And a.ҽ��״̬ In (3, 8) " & vbNewLine & _
            " And ����ʱ�� Between [4] and [5] "

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, mintBaby, datBegin, datEnd)
    
    udtDataItem = mudtTimeLine.colFooterData(M_CON_KEY_����)
    Set colFoot = New Collection
    
    For i = 1 To rsTmp.RecordCount
        With udtDataInfo
            .Value = rsTmp!���� & ""
            .Time = Format(rsTmp!����ʱ�� & "", "YYYY-MM-DDTHH:MM:SS")
            .RowIndex = rsTmp!ID
            .Tag = rsTmp!ҽ������ & ""
        End With
        colFoot.Add udtDataInfo, "_" & (colFoot.Count + 1)
        rsTmp.MoveNext
    Next
    
    Set udtDataItem.ListData = colFoot
    mudtTimeLine.colFooterData.Remove (M_CON_KEY_����)
    mudtTimeLine.colFooterData.Add udtDataItem, M_CON_KEY_����, M_CON_KEY_����
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncLoadOperationAfter()
'����:������������
    Dim colHead As Collection
    Dim udtDataInfo As DataInfo
    Dim udtDataItem As DataItem
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim datBegin As Date, datEnd As Date
    '���ڷ�ҳ����,��Ҫ����ʼʱ�����ó���Ժʱ��ͳ�Ժʱ�䣬���򵱷�ҳʱ���û��������¼ʱ,��������ʾ����������
    datBegin = CDate(Format(mDatIn, "YYYY-MM-DD 00:00:00"))
    datEnd = CDate(Format(mdatOut, "YYYY-MM-DD 23:59:59"))
    
    '������ҽ��
    strSQL = "Select a.����ʱ�� " & vbNewLine & _
            "From ����ҽ����¼ A " & vbNewLine & _
            "Where a.����id = [1] And a.��ҳid = [2] And NVL(a.Ӥ��,0) =[3] And a.������� = 'F' And Nvl(a.���id, 0) = 0 And a.ҽ��״̬ In (3, 8) " & vbNewLine & _
            " And ����ʱ�� Between [4] and [5] "

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, mintBaby, datBegin, datEnd)
    
    udtDataItem = mudtTimeLine.colHeaderData(M_CON_KEY_����������)
    Set colHead = New Collection
    For i = 1 To rsTmp.RecordCount
        With udtDataInfo
            .Time = Format(rsTmp!����ʱ�� & "", "YYYY-MM-DDTHH:MM:SS")
            .TimeEnd = Format(mdatOut, "YYYY-MM-DDTHH:MM:SS")      '
        End With
        colHead.Add udtDataInfo, "_" & (colHead.Count + 1)
        rsTmp.MoveNext
    Next
    
    Set udtDataItem.ListData = colHead
    mudtTimeLine.colHeaderData.Remove (M_CON_KEY_����������)
    mudtTimeLine.colHeaderData.Add udtDataItem, M_CON_KEY_����������, , M_CON_KEY_סԺ����
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncLoadAdvice()
'����:��������ҽ��
    Dim colFoot As Collection
    Dim udtDataInfo As DataInfo
    Dim udtDataItem As DataItem
    Dim strSQL As String, strBegin As String, strEnd As String
    Dim rsTmp As ADODB.Recordset
    Dim rsDrug As ADODB.Recordset
    Dim i As Long, j As Long, lngTemp As Long
    Dim lngRow As Long, lngDay As Long
    Dim lngColor As Long
    Dim datBegin As Date, datEnd As Date, datCurr As Date, datTemp As Date
    
    '��������ҽ��:��ʱҽ���ͳ���ҽ���ֱ����
    datBegin = CDate(Format(mDatBegin, "YYYY-MM-DD 00:00:00"))
    datEnd = CDate(Format(mDatEnd, "YYYY-MM-DD 23:59:59"))
    datCurr = zlDatabase.Currentdate
    strSQL = "Select a.ID,a.ҽ������, b.����,a.��ʼִ��ʱ��,a.ִ����ֹʱ��,a.ҽ����Ч,������� " & vbNewLine & _
        "From ����ҽ����¼ A, ������ĿĿ¼ B" & vbNewLine & _
        "Where a.������Ŀid = b.Id And a.����id = [1] And a.��ҳid = [2] And NVL(a.Ӥ��,0) =[3]  And Not a.������� In ('G', 'F', 'D', 'C', '5', '6', '7') And" & vbNewLine & _
        "      Not (Nvl(b.��������, 0) In ('2', '3', '4', '6', '8') And a.������� = 'E') And Nvl(���id, 0) = 0 And" & vbNewLine & _
        "      (a.��ʼִ��ʱ�� Between [4] And [5] OR (a.ҽ����Ч=0 And a.��ʼִ��ʱ�� < [4] And NVL(a.ִ����ֹʱ��,[5])>[4]))  And" & vbNewLine & _
        "      ((a.ҽ����Ч = 1 And a.ҽ��״̬ In (3, 8)) Or (a.ҽ����Ч = 0 And a.ҽ��״̬ In (3, 5, 6, 7, 8, 9)))" & vbNewLine & _
        "Order By a.��ʼִ��ʱ��, a.���"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, mintBaby, datBegin, datEnd)
    
    rsTmp.Filter = "ҽ����Ч = 1"
    udtDataItem = mudtTimeLine.colFooterData(M_CON_KEY_��������)
    Set colFoot = New Collection
    
    For i = 1 To rsTmp.RecordCount
        Select Case UCase(rsTmp!������� & "")
        Case "H"
            lngColor = COLOR_ǳ��
        Case "L"
            lngColor = COLOR_ǳ��
        Case "Z"
            lngColor = COLOR_����
        Case "I"  '��ʳ
            lngColor = COLOR_ǳ��
        Case "K"   '��Ѫ
            lngColor = COLOR_ǳ��
        Case "M", "4"   '����,����
            lngColor = COLOR_ǳ��
        Case "E"
            lngColor = COLOR_����
        Case Else
            lngColor = vbWhite
        End Select
        With udtDataInfo
            .Value = rsTmp!���� & ""
            .Time = Format(rsTmp!��ʼִ��ʱ�� & "", "YYYY-MM-DDTHH:MM:SS")
            .BackgroundColor = FuncColorRGB(lngColor)
            .RowIndex = rsTmp!ID
            .Tip = "ҽ������:" & rsTmp!���� & "" & vbCrLf & _
                   "��Чʱ��:" & Format(rsTmp!��ʼִ��ʱ�� & "", "YYYY-MM-DDTHH:MM:SS")
        End With
        colFoot.Add udtDataInfo, "_" & (colFoot.Count + 1)
        rsTmp.MoveNext
    Next
    
    Set udtDataItem.ListData = colFoot
    mudtTimeLine.colFooterData.Remove (M_CON_KEY_��������)
    mudtTimeLine.colFooterData.Add udtDataItem, M_CON_KEY_��������, M_CON_KEY_��������
    
    rsTmp.Filter = "ҽ����Ч = 0"
    udtDataItem = mudtTimeLine.colFooterData(M_CON_KEY_��������)
    Set colFoot = New Collection
    Set rsDrug = InitRS
    For i = 1 To rsTmp.RecordCount
        If Between(CDate(rsTmp!��ʼִ��ʱ�� & ""), datBegin, datEnd) Then
            strBegin = Format(rsTmp!��ʼִ��ʱ��, "YYYY-MM-DD ")
        Else
            strBegin = Format(datBegin, "YYYY-MM-DD ")
        End If
        If Nvl(rsTmp!ִ����ֹʱ��) = "" Then
            If DateDiff("D", datEnd, datCurr) >= 0 Then
                datTemp = datEnd
            Else
                datTemp = datCurr
            End If
        Else
            If CDate(rsTmp!ִ����ֹʱ��) > datEnd Then
                datTemp = datEnd
            Else
                datTemp = CDate(rsTmp!ִ����ֹʱ��)
            End If
        End If
        strEnd = Format(datTemp, "YYYY-MM-DD 23:59:59")
        Select Case UCase(rsTmp!������� & "")
        Case "H"
            lngColor = COLOR_ǳ��
        Case "L"
            lngColor = COLOR_ǳ��
        Case "Z"
            lngColor = COLOR_����
        Case "I"  '��ʳ
            lngColor = COLOR_ǳ��
        Case "K"   '��Ѫ
            lngColor = COLOR_ǳ��
            lngColor = &HFF
        Case "M"    '����
            lngColor = COLOR_ǳ��
        Case "E"
            lngColor = COLOR_����
        Case Else
            lngColor = vbWhite
        End Select
        lngTemp = 0
        For j = 1 To lngRow
            rsDrug.Filter = "�к�=" & j & " And ���� >= '" & Format(strBegin, "YYYY-MM-DD") & "'"
            If rsDrug.RecordCount = 0 Then
                lngTemp = j
                Exit For
            End If
        Next
        If j > lngRow Then lngRow = lngRow + 1: lngTemp = lngRow
        If strEnd = "" Then
            lngDay = 1
        Else
            lngDay = DateDiff("D", Format(strBegin, "YYYY-MM-DD"), Format(strEnd, "YYYY-MM-DD"))
            If lngDay < 1 Then lngDay = 1
        End If
        rsDrug.AddNew
        For j = 0 To lngDay
            rsDrug!�к� = lngTemp
            rsDrug!���� = Format(DateAdd("D", j, Format(strBegin, "YYYY-MM-DD")), "YYYY-MM-DD")
        Next
        rsDrug.UpdateBatch
    
        With udtDataInfo
            .RowNumber = lngTemp
            .Value = rsTmp!���� & ""
            .Time = Format(strBegin, "YYYY-MM-DDTHH:MM:SS")
            .TimeEnd = Format(strEnd, "YYYY-MM-DDTHH:MM:SS")
            .BackgroundColor = FuncColorRGB(lngColor)
            .Tip = "ҽ������:" & rsTmp!ҽ������ & vbCrLf & _
                   "��Чʱ��:" & Format(rsTmp!��ʼִ��ʱ�� & "", "YYYY-MM-DD HH:MM:SS") & IIf(rsTmp!ִ����ֹʱ�� & "" <> "", vbCrLf & "��ֹʱ��:" & Format(rsTmp!ִ����ֹʱ�� & "", "YYYY-MM-DD HH:MM:SS"), "")
        End With
        colFoot.Add udtDataInfo, "_" & (colFoot.Count + 1)
        rsTmp.MoveNext
    Next
    
    Set udtDataItem.ListData = colFoot
    mudtTimeLine.colFooterData.Remove (M_CON_KEY_��������)
    mudtTimeLine.colFooterData.Add udtDataItem, M_CON_KEY_��������, M_CON_KEY_����
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncLoadItemNurse()
'����:���ػ�����Ŀ
    Dim colTmp As Collection
    Dim udtDataInfo As DataInfo
    Dim udtDataItem As DataItem
    Dim strSQL As String, strKey  As String, strTemp As String
    Dim strRecordID As String
    Dim rsTmp As ADODB.Recordset, rsCopy As ADODB.Recordset
    Dim arrTag As Variant
    Dim i As Long, j As Long
    Dim lngRow As Long
    Dim datBegin As Date, datEnd As Date

    datBegin = CDate(Format(mDatBegin, "YYYY-MM-DD 00:00:00"))
    datEnd = CDate(Format(mDatEnd, "YYYY-MM-DD 23:59:59"))
    
    '���ػ�����Ŀ
    strSQL = "Select c.����ʱ��,d.��¼ID, d.��Ŀ���, d.��¼����, d.��¼���,d.���²�λ, f.��¼��, f.��¼��,NVL(f.��¼Ƶ��,2) as ��¼Ƶ�� " & vbNewLine & _
            "From ���˻����ļ� A, �����ļ��б� B, ���˻������� C, ���˻�����ϸ D, ���¼�¼��Ŀ F" & vbNewLine & _
            "Where a.��ʽid = b.Id And a.Id = c.�ļ�id And c.Id = d.��¼id And d.��Ŀ��� = f.��Ŀ��� And a.����id = [1] And a.��ҳid = [2] And" & vbNewLine & _
            "      NVL(a.Ӥ��,0) = [3] And b.���� = 3 And b.���� = -1 And c.����ʱ�� Between [4] And [5]" & vbNewLine & _
            "Order By f.��¼��, d.��Ŀ���, c.����ʱ��,d.��¼��� "
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, mintBaby, datBegin, datEnd)
    '����,����,����
    lngRow = mudtTimeLine.colMeasureData.Count
    For i = 1 To lngRow
        udtDataItem = mudtTimeLine.colMeasureData(i)
        strKey = udtDataItem.ItemTag
        Set colTmp = New Collection
        If strKey = "K_1" Then
            '���£�37/38  �����±�ʾ��,ͬһʱ��㣨ͬһ��¼ID��,������������
            rsTmp.Filter = "��Ŀ���=1"
            strRecordID = ""
            Call FuncClearUDT(udtDataInfo)
            For j = 1 To rsTmp.RecordCount
                If j = 1 Then strRecordID = rsTmp!��¼ID
                If strRecordID <> rsTmp!��¼ID Then
                    strRecordID = rsTmp!��¼ID
                    colTmp.Add udtDataInfo, "_" & (colTmp.Count + 1)
                    Call FuncClearUDT(udtDataInfo)
                End If
                With udtDataInfo
                    If rsTmp!��¼��� = 0 Then
                        .Time = Format(rsTmp!����ʱ�� & "", "YYYY-MM-DDTHH:MM:SS")
                        .NumberValue = IIf(Val(rsTmp!��¼���� & "") = 0, "", Val(rsTmp!��¼���� & ""))
                        arrTag = Split(rsTmp!��¼�� & "", ",")
                        If rsTmp!���²�λ = "����" Then
                            .LegendType = arrTag(0)
                        ElseIf rsTmp!���²�λ = "Ҹ��" Then
                            .LegendType = arrTag(1)
                        ElseIf rsTmp!���²�λ = "����" Then
                            .LegendType = arrTag(2)
                        ElseIf rsTmp!���²�λ = "����" Then
                            .LegendType = arrTag(3)
                        ElseIf rsTmp!���²�λ = "����" Then
                            .LegendType = arrTag(4)
                        End If
                    ElseIf rsTmp!��¼��� = 1 Then
                        .BalloonValue = rsTmp!��¼���� & ""
                        .BalloonLegendType = "����Բ"
                    End If
                End With
                If rsTmp.RecordCount = j Then
                    colTmp.Add udtDataInfo, "_" & (colTmp.Count + 1)
                    Call FuncClearUDT(udtDataInfo)
                End If
                rsTmp.MoveNext
            Next
        ElseIf strKey = "K_2" Then
            '����
            rsTmp.Filter = "��Ŀ��� =-1"
            Set rsCopy = zlDatabase.CopyNewRec(rsTmp)
            rsTmp.Filter = "��Ŀ��� = 2"
            rsTmp.Sort = "����ʱ��"
            Do While Not rsTmp.EOF
                Call FuncClearUDT(udtDataInfo)
                With udtDataInfo
                    .Time = Format(rsTmp!����ʱ�� & "", "YYYY-MM-DDTHH:MM:SS")
                    .NumberValue = rsTmp!��¼���� & ""
                    arrTag = Split(rsTmp!��¼�� & "", ",")
                    '����ȱʡ�ǡ�H��,ȱʡȡ����ֵ
                    If rsTmp!���²�λ = "����" Then
                        .LegendType = "H"     'H
                    Else
                        .LegendType = arrTag(0) '
                    End If
                    
                    rsCopy.Filter = "����ʱ�� =#" & rsTmp!����ʱ�� & "#"
                    If Not rsCopy.EOF Then
                        .ShadowValue = rsCopy!��¼���� & ""
                        .ShadowLegendType = rsCopy!��¼�� & ""
                    Else
                        .ShadowValue = rsTmp!��¼���� & ""
                        .ShadowLegendType = "��"
                    End If
                End With
                colTmp.Add udtDataInfo, "_" & (colTmp.Count + 1)
                rsTmp.MoveNext
            Loop
        ElseIf strKey = "K_-1" Then
            rsTmp.Filter = "��Ŀ��� =-1"
            rsTmp.Sort = "����ʱ��"
            Call FuncClearUDT(udtDataInfo)
            For j = 1 To rsTmp.RecordCount
                With udtDataInfo
                    .Time = Format(rsTmp!����ʱ�� & "", "YYYY-MM-DDTHH:MM:SS")
                    .NumberValue = rsTmp!��¼���� & ""
                    .LegendType = rsTmp!��¼�� & ""
                End With
                colTmp.Add udtDataInfo, "_" & (colTmp.Count + 1)
                Call FuncClearUDT(udtDataInfo)
                rsTmp.MoveNext
            Next
        End If
        If i < lngRow Then strTemp = mudtTimeLine.colMeasureData(i + 1).ItemTag
        Set udtDataItem.ListData = colTmp
        mudtTimeLine.colMeasureData.Remove (strKey)
        If i < lngRow Then
            mudtTimeLine.colMeasureData.Add udtDataItem, strKey, strTemp
        Else
            mudtTimeLine.colMeasureData.Add udtDataItem, strKey
        End If
    Next
    
    'Ѫѹ
    Call FuncGetItemNurse(datBegin, datEnd, rsTmp)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function FuncColorRGB(lngColor As Long) As String
'����:����RGB��ɫֵ���ַ���
    Dim Color(2) As Byte
    Color(0) = (lngColor Mod 256)
    Color(1) = ((lngColor Mod 65536) \ 256)
    Color(2) = (lngColor \ 65536)
    FuncColorRGB = Color(0) & "," & Color(1) & "," & Color(2)
End Function

Public Sub SetFontSize(ByVal bytSize As Byte)
'����:����ҽ���嵥�������С
'���:bytSize��0-С(ȱʡ)��1-��
    Dim strDesignInfo As String
    Dim strData As String
    
    mbytFont = bytSize

    Call zlControl.SetPubFontSize(Me, bytSize)

    strDesignInfo = FuncMakeXMLDesign(mudtDesign)
    LogWrite "סԺһ���ĵ�����־", "" & glngModul, "SetFontSize", "Design_FontSize:" & vbCrLf & strDesignInfo
    strData = FuncMakeXMLTimeLine(mudtTimeLine)
    LogWrite "סԺһ���ĵ�����־", "" & glngModul, "SetFontSize", "Data_FontSize:" & vbCrLf & strData
    mtimeLineControl.UpdateDesignInfo strDesignInfo
    mtimeLineControl.UpdateData strData
    mtimeLineControl.RefreshAll
    If mblnMeasureArea Then
        mtimeLineControl.ShowMeasureArea
    Else
        mtimeLineControl.HideMeasureArea
    End If
End Sub

Public Sub zlExecuteCommandBars()
'����:סԺҽ��վ����
    Dim objContrl As CommandBarControl
    Set objContrl = cbsSub.FindControl(xtpControlButton, conMenu_Process_Zoom)
    objContrl.Execute
End Sub

Private Sub mtimeLineControl_DataMouseDoubleClick(ByVal sender As Variant, ByVal e As ZLSoft_BusinessHome_ClientControl_TimeLineBase.IEventArgsDataInfo)
'�¼��������ڱ������˫�����쳣
End Sub

Private Sub mtimeLineControl_DataTitleMouseDoubleClick(ByVal sender As Variant, ByVal e As ZLSoft_BusinessHome_ClientControl_TimeLineBase.IEventArgsDataInfo)
'�¼��������ڱ������˫�����쳣
End Sub

Private Sub mtimeLineControl_DateMouseDoubleClick(ByVal sender As Variant, ByVal e As ZLSoft_BusinessHome_ClientControl_TimeLineBase.IEventArgsDateInfo)
'�¼��������ڱ������˫�����쳣
End Sub

Private Sub mtimeLineControl_TimeLineMouseClick(ByVal sender As Variant, ByVal e As ZLSoft_BusinessHome_ClientControl_TimeLineBase.IEventArgsTimeLineMouse)
'�¼��������ڱ�����굥�����쳣
End Sub

Private Function FuncCompareTime(ByVal strBegin As String, ByVal strEnd As String, ByVal strCheckVal As String) As Long
'����:���ص�ǰʱ�����м�ʱ��(��ʼʱ�����ֹʱ����м�ʱ��)����������
'����:
'strBegin-��ʼʱ��
'strEnd-��ֹʱ��
'strCheckVal-Ҫ�Աȵ�ʱ��

    Dim lngDiff As Double
    Dim datTemp As Date
    
    lngDiff = DateDiff("s", CDate(strBegin), CDate(strEnd)) \ 2
    datTemp = DateAdd("s", lngDiff, CDate(strBegin))

    FuncCompareTime = Abs(DateDiff("s", datTemp, CDate(strCheckVal)))
End Function


Private Function FuncGetNodeSN(ByVal strSN As String) As String
'����:���ػ��ܽڵ����������
    Dim strRet As String
    Dim lngPos As Long
    
    mrs������Ŀ.Filter = "�����=" & strSN
    
    If mrs������Ŀ.RecordCount = 0 Then
        strRet = strSN
    Else
        Do While Not mrs������Ŀ.EOF
            lngPos = mrs������Ŀ.AbsolutePosition
            strRet = strRet & "," & FuncGetNodeSN(mrs������Ŀ!��� & "")
            mrs������Ŀ.Filter = "�����=" & strSN
            mrs������Ŀ.AbsolutePosition = lngPos
            mrs������Ŀ.MoveNext
        Loop
        strRet = Mid(strRet, 2)
    End If
    If InStr("," & strRet & ",", "," & strSN & ",") = 0 Then strRet = strSN & "," & strRet
    FuncGetNodeSN = strRet
End Function

Private Function InitRS(Optional ByVal bytFunc As Byte = 1) As ADODB.Recordset
'����:�����¼��
    Dim rs As ADODB.Recordset
    Dim strFields As String
    Dim strFieldName As String
    Dim lngLen As Long
    Dim FieldType As DataTypeEnum
    Dim i As Long, j As Long
    
    Dim arrField As Variant
    Dim arrSubFeld As Variant '�ֶ�����|�ֶ�����|�ֶγ��� ȱʡ�ֶ����� ΪadVarChar
    
    If bytFunc = 1 Then
        strFields = "�к�|adBigInt|18,����|adVarChar|10"
    Else
        strFields = "StartDate|adVarChar|20,TickWidth|adVarChar|20"
    End If
    
    Set rs = New ADODB.Recordset
    '-----------------------------------------
    With rs.Fields
        arrField = Split(strFields, ",")
        For i = LBound(arrField) To UBound(arrField)
            arrSubFeld = Split(arrField(i), "|")
            strFieldName = arrSubFeld(0)
            Select Case UCase(arrSubFeld(1) & "")
            Case UCase("adVarChar")
                FieldType = adVarChar
            Case UCase("adBigInt")
                FieldType = adBigInt
            End Select
            lngLen = Val(arrSubFeld(2))
            .Append strFieldName, FieldType, lngLen
        Next
    End With
    '---------------------------------------
    rs.CursorLocation = adUseClient
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenStatic
    rs.Open
    '----------------------------------
    Set InitRS = rs
End Function

Private Function FuncGetSubString(ByVal strInput As String, ByVal lngWidth As Long, Optional ByVal lngStart As Long = 9) As String
'����:���ݴ���ĳ��Ƚ�ȡ�ַ�
    Dim strRet As String
    Dim i As Long
    
    If Me.TextWidth(strInput) > lngWidth Then
        strRet = Mid(strInput, 1, lngStart)
        For i = lngStart + 1 To Len(strInput)
            If Me.TextWidth(strRet & Mid(strInput, i, 1) & "...") >= lngWidth Then
                strRet = strRet & "..."
                Exit For
            Else
                strRet = strRet & Mid(strInput, i, 1)
            End If
        Next
    Else
        strRet = strInput
    End If
    FuncGetSubString = strRet
End Function

Private Sub FuncGetItemNurse(ByVal datBegin As Date, ByVal datEnd As Date, ByVal rsItem As ADODB.Recordset)
'����:
    Dim i As Long, j As Long, k As Long
    Dim lngRow As Long
    Dim lngPos As Long, lngPosOne As Long
    Dim lngMinDiff As Long, lngDateDiff As Long

    Dim bytItem As Byte '0-��ͨ��Ŀ;1-������Ŀ;2-������Ŀ,
    Dim bytSplitNum As Byte
    Dim datCurr As Date
    
    Dim colTmp As Collection, colblood As Collection
    Dim udtDataInfo As DataInfo
    Dim udtDataItem As DataItem
    Dim arrTmp As Variant, varItem As Variant
    Dim strFreB As String, strFreE As String, strMin As String, strMax As String
    Dim strBegin As String, strEnd As String
    Dim strTemp As String, strItem As String, strPrveTime As String
    Dim rsCopy As ADODB.Recordset
    Dim blnDo As Boolean
    
    On Error GoTo errH
    lngRow = mudtTimeLine.colFooterData("K_������Ŀ").ListData.Count
    For i = 1 To lngRow
        udtDataItem = mudtTimeLine.colFooterData("K_������Ŀ").ListData(i)
        Set colTmp = New Collection
        arrTmp = Split(udtDataItem.ItemTag, ",")   '����,��Ŀ���,����,����,�����
        bytItem = Val(arrTmp(2))
        
        If arrTmp(0) = "K_3" Then  '����
            rsItem.Filter = "��Ŀ��� =3"
            For j = 1 To rsItem.RecordCount
                With udtDataInfo
                    .Value = rsItem!��¼���� & ""
                    .Time = Format(rsItem!����ʱ�� & "", "YYYY-MM-DDTHH:MM:SS")
                End With
                colTmp.Add udtDataInfo, "_" & (colTmp.Count + 1)
                rsItem.MoveNext
            Next
        Else
            If arrTmp(1) = "5" Then
                Set colblood = New Collection
                arrTmp(1) = "4"  '����ѹ  �Ƚ�����ѹ����
            End If
LineBlood:
            If bytItem = 2 Then
                strTemp = FuncGetNodeSN(arrTmp(1))
                varItem = Split(strTemp, ","): strItem = ""
                For j = LBound(varItem) To UBound(varItem)
                    strItem = strItem & " OR ��Ŀ��� =" & varItem(j)
                Next
                rsItem.Filter = Mid(strItem, 4)
            Else
                rsItem.Filter = "��Ŀ��� = " & arrTmp(1)
            End If
            rsItem.Sort = "����ʱ��,��Ŀ���"
            If rsItem.RecordCount > 0 Then
                Set rsCopy = zlDatabase.CopyNewRec(rsItem)  '�������Ŀ��ҳ��������
                'Ƶ��
                bytSplitNum = Val(rsCopy!��¼Ƶ�� & "")
                If bytItem = 0 Then
                    mrsFrequency.Filter = "Ƶ��=" & bytSplitNum
                    If mrsFrequency.RecordCount > 0 Then
                        mrsFrequency.MoveFirst
                        strBegin = mrsFrequency!��ʼ & ""
                        mrsFrequency.MoveLast
                        strEnd = mrsFrequency!���� & ""
                    Else
                        strBegin = "00:00"
                        strFreE = "23:59"
                    End If
                ElseIf bytItem = 1 Or bytItem = 2 Then
                    If bytSplitNum = 1 Then
                        mrs����ʱ��.Filter = "��� = 3"
                    Else
                        mrs����ʱ��.Filter = "��� < 3"
                    End If
                    If mrs����ʱ��.RecordCount > 0 Then
                        mrs����ʱ��.MoveFirst
                        strBegin = mrs����ʱ��!��ʼ & ""
                        mrs����ʱ��.MoveLast
                        strEnd = mrs����ʱ��!���� & ""
                    Else
                        strBegin = "00:00"
                        strEnd = "23:59"
                    End If
                End If
                datCurr = datBegin
                Do While datCurr <= datEnd
                    If strBegin <= strEnd Then
                        rsCopy.Filter = "����ʱ�� >= #" & Format(datCurr, "YYYY-MM-DD ") & Format(strBegin, "HH:MM:SS") & "# And ����ʱ�� <= #" & Format(datCurr, "YYYY-MM-DD ") & Format(strEnd, "HH:MM:SS") & "#"
                    Else
                        rsCopy.Filter = "����ʱ�� >= #" & Format(datCurr, "YYYY-MM-DD ") & Format(strBegin, "HH:MM:SS") & "# And ����ʱ�� <= #" & Format(DateAdd("d", 1, datCurr), "YYYY-MM-DD ") & Format(strEnd, "HH:MM:SS") & "#"
                    End If
                    
                    If rsCopy.RecordCount > 0 Then
                        rsCopy.Sort = "����ʱ��,��Ŀ���"
                        strItem = ""
                        For k = 1 To bytSplitNum - 1
                            strItem = strItem & ","
                        Next
                        If strItem = "" Then
                            varItem = Array("")
                        Else
                            varItem = Split(strItem, ",")
                        End If
                    End If
                    
                    strItem = "": strMin = "": strMax = ""
                    Do While Not rsCopy.EOF
                        If bytItem = 0 Then
                            mrsFrequency.Filter = "Ƶ��=" & bytSplitNum
                            lngPos = rsCopy.AbsolutePosition
                            For k = 1 To mrsFrequency.RecordCount
                                If mrsFrequency!��ʼ & "" <= mrsFrequency!���� & "" Then
                                    strFreB = Format(Format(datCurr, "YYYY-MM-DD ") & mrsFrequency!��ʼ, "YYYY-MM-DD HH:MM:SS")
                                    strFreE = Format(Format(datCurr, "YYYY-MM-DD ") & mrsFrequency!����, "YYYY-MM-DD HH:MM:SS")
                                Else
                                    strFreB = Format(Format(datCurr, "YYYY-MM-DD ") & mrsFrequency!��ʼ, "YYYY-MM-DD HH:MM:SS")
                                    strFreE = Format(Format(DateAdd("d", 1, datCurr), "YYYY-MM-DD ") & mrsFrequency!����, "YYYY-MM-DD HH:MM:SS")
                                End If
                                If Between(CDate(rsCopy!����ʱ��), CDate(strFreB), CDate(strFreE)) Then
                                    If varItem(mrsFrequency!��� - 1) <> "" Then Exit For
                                    If mrsFrequency!��� = 1 Then
                                        'ȡ��һ��
                                        varItem(mrsFrequency!��� - 1) = rsCopy!��¼����
                                        Exit For
                                    ElseIf mrsFrequency!��� = 2 Then
                                        'ȡ�м�һ��
                                        lngPosOne = lngPos
                                        lngMinDiff = -1
                                        Do While Not rsCopy.EOF
                                            If Between(CDate(rsCopy!����ʱ��), CDate(strFreB), CDate(strFreE)) Then
                                                lngDateDiff = FuncCompareTime(strFreB, strFreE, Format(rsCopy!����ʱ�� & "", "YYYY-MM-DD HH:MM:SS"))
                                                If lngMinDiff >= lngDateDiff Or lngMinDiff = -1 Then
                                                    lngMinDiff = lngDateDiff
                                                    lngPosOne = rsCopy.AbsolutePosition
                                                Else
                                                    Exit Do
                                                End If
                                            Else
                                                Exit Do
                                            End If
                                            rsCopy.MoveNext
                                        Loop
                                        rsCopy.AbsolutePosition = lngPosOne
                                        varItem(mrsFrequency!��� - 1) = rsCopy!��¼����
                                        Exit For
                                    ElseIf mrsFrequency!��� = 3 Then
                                        'ȡ���һ��
                                        rsCopy.MoveNext
                                        If rsCopy.EOF Then
                                            rsCopy.AbsolutePosition = lngPos
                                            varItem(mrsFrequency!��� - 1) = rsCopy!��¼����
                                        ElseIf Format(rsCopy!����ʱ�� & "", "YYYY-MM-DD HH:MM:SS") > strFreE Then
                                            rsCopy.AbsolutePosition = lngPos
                                            varItem(mrsFrequency!��� - 1) = rsCopy!��¼����
                                        Else
                                            rsCopy.AbsolutePosition = lngPos
                                        End If
                                        Exit For
                                    End If
                                End If
                                mrsFrequency.MoveNext
                            Next
                        ElseIf bytItem = 1 Or bytItem = 2 Then
                            '������Ŀ��������Ŀ
                            If bytSplitNum = 1 Then
                                mrs����ʱ��.Filter = "��� = 3"
                            Else
                                mrs����ʱ��.Filter = "��� < 3"
                            End If
                            If bytItem = 1 Then
                                For k = 1 To mrs����ʱ��.RecordCount
                                    If mrs����ʱ��!��ʼ & "" <= mrs����ʱ��!���� & "" Then
                                        strFreB = Format(Format(datCurr, "YYYY-MM-DD ") & mrs����ʱ��!��ʼ, "YYYY-MM-DD HH:MM:SS")
                                        strFreE = Format(Format(datCurr, "YYYY-MM-DD ") & mrs����ʱ��!����, "YYYY-MM-DD HH:MM:SS")
                                    Else
                                        strFreB = Format(Format(datCurr, "YYYY-MM-DD ") & mrs����ʱ��!��ʼ, "YYYY-MM-DD HH:MM:SS")
                                        strFreE = Format(Format(DateAdd("d", 1, datCurr), "YYYY-MM-DD ") & mrs����ʱ��!����, "YYYY-MM-DD HH:MM:SS")
                                    End If
                                    If varItem(k - 1) = "" Then
                                        If Between(Format(rsCopy!����ʱ�� & "", "YYYY-MM-DD HH:MM:SS"), strFreB, strFreE) Then
                                            If Val(strMax) < Val(rsCopy!��¼���� & "") Or strMax = "" Then strMax = rsCopy!��¼���� & ""
                                            If Val(strMin) > Val(rsCopy!��¼���� & "") Or strMin = "" Then strMin = rsCopy!��¼���� & ""
                                            Exit For
                                        End If
                                        If (strMin <> "" Or strMax <> "") Then
                                            If strMin = strMax Then
                                                varItem(k - 1) = strMin
                                            ElseIf strMin <> strMax Then
                                                varItem(k - 1) = strMin & "-" & strMax
                                            End If
                                            strMin = "": strMax = ""
                                        End If
                                    End If
                                    mrs����ʱ��.MoveNext
                                Next
                            ElseIf bytItem = 2 Then
                                '������Ŀ
                                For k = 1 To mrs����ʱ��.RecordCount
                                    If mrs����ʱ��!��ʼ & "" <= mrs����ʱ��!���� & "" Then
                                        strFreB = Format(Format(datCurr, "YYYY-MM-DD ") & mrs����ʱ��!��ʼ, "YYYY-MM-DD HH:MM:SS")
                                        strFreE = Format(Format(datCurr, "YYYY-MM-DD ") & mrs����ʱ��!����, "YYYY-MM-DD HH:MM:SS")
                                    Else
                                        strFreB = Format(Format(datCurr, "YYYY-MM-DD ") & mrs����ʱ��!��ʼ, "YYYY-MM-DD HH:MM:SS")
                                        strFreE = Format(Format(DateAdd("d", 1, datCurr), "YYYY-MM-DD ") & mrs����ʱ��!����, "YYYY-MM-DD HH:MM:SS")
                                    End If
                                    If varItem(k - 1) = "" Then
                                          If Between(Format(rsCopy!����ʱ�� & "", "YYYY-MM-DD HH:MM:SS"), strFreB, strFreE) Then
                                            strMax = Val(strMax) + Val(rsCopy!��¼���� & "")
                                            Exit For
                                        End If
                                        If strMax <> "" Then
                                            varItem(k - 1) = strMax
                                            strMax = ""
                                        End If
                                    End If
                                    mrs����ʱ��.MoveNext
                                Next
                             End If
                        End If
                        rsCopy.MoveNext
                    Loop
                    '�����커�����ݼ��뵽���ݼ���
                    If rsCopy.RecordCount > 0 Then
                        If bytItem = 1 And (strMin <> "" Or strMax <> "") Then
                            '(strMin <> "" Or strMax <> "") ��������������ʱ�䲻�ڻ���ʱ��,����Kֵ����3�����±�Խ��
                            If varItem(k - 1) = "" Then
                                If strMin = strMax Then
                                    varItem(k - 1) = strMin
                                ElseIf strMin <> strMax Then
                                    varItem(k - 1) = strMin & "-" & strMax
                                End If
                                strMin = "": strMax = ""
                            End If
                            If Val(mstr��ʾ����) = 0 Then
                                strPrveTime = DateAdd("d", 1, datCurr)
                            End If
                        ElseIf bytItem = 2 And strMax <> "" Then
                            If varItem(k - 1) = "" Then
                                varItem(k - 1) = strMax
                                strMax = ""
                            End If
                            If Val(mstr��ʾ����) = 0 Then
                                strPrveTime = DateAdd("d", 1, datCurr)
                            End If
                        End If
                        
                        strItem = "": If arrTmp(1) = "5" Then strTemp = colblood(Format(datCurr, "YYYY-MM-DDTHH:MM:SS"))
                        For k = LBound(varItem) To UBound(varItem)
                            If arrTmp(1) = "5" Then
                                strItem = strItem & "," & IIf(Split(strTemp, ",")(k) <> "" Or varItem(k) <> "", Split(strTemp, ",")(k) & "/" & varItem(k), Split(strTemp, ",")(k) & varItem(k))
                            Else
                                strItem = strItem & "," & varItem(k)
                            End If
                        Next
                        strItem = Mid(strItem, 2)
                        If arrTmp(1) = "4" Then
                            colblood.Add strItem, Format(datCurr, "YYYY-MM-DDTHH:MM:SS")
                        Else
                            With udtDataInfo
                                .Value = strItem
                                .Time = Format(datCurr, "YYYY-MM-DDTHH:MM:SS")
                                .Tag = strItem
                                colTmp.Add udtDataInfo, "_" & (colTmp.Count + 1)
                            End With
                        End If
                    End If
                    datCurr = DateAdd("d", 1, datCurr)
                Loop
            End If
        End If
        
        If arrTmp(1) = "4" Then arrTmp(1) = "5": GoTo LineBlood 'Ѫѹ:����ѹ/����ѹ
        
        If i < lngRow Then strTemp = mudtTimeLine.colFooterData("K_������Ŀ").ListData(i + 1).ItemTag
        Set udtDataItem.ListData = colTmp
        mudtTimeLine.colFooterData("K_������Ŀ").ListData.Remove (arrTmp(0))
        If i < lngRow Then
            mudtTimeLine.colFooterData("K_������Ŀ").ListData.Add udtDataItem, arrTmp(0), Split(strTemp, ",")(0)
        Else
            mudtTimeLine.colFooterData("K_������Ŀ").ListData.Add udtDataItem, arrTmp(0)
        End If
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

