VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmHistoryStudy 
   BorderStyle     =   0  'None
   ClientHeight    =   4980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex8Ctl.VSFlexGrid vsfStudy 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   8535
      _cx             =   15055
      _cy             =   7435
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
      BackColorSel    =   16761024
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
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
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
      ExplorerBar     =   3
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
   Begin XtremeCommandBars.ImageManager imgList 
      Left            =   600
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmHistoryStudy.frx":0000
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   120
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmHistoryStudy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngModule As Long
Private mblnMoved As Boolean
Private mlngRow As Long
Private mdtBegin As Date
Private mdtEnd As Date
Private mblnCustom As Boolean
Private mlngPatId As Long
Private mlngCur����ID As Long
Private mlngLinkID As Long
Private mstrCanUse����IDs As String
Private mblnAllDepts As Boolean
Private mblnRelatingPatient As Boolean
Private mLngAdvice As Long
Private mblnDocPatient As Boolean        '�Ƿ�סԺ����
Private mlngBabyNum As Long              'Ӥ�����
Private mblnImageEnable As Boolean
Private mblnReportEnable As Boolean
Private mTPListCfg As TPListCfg
Private mblnListCfgOk As Boolean

Private Enum FilterID
    ID_������� = 1
    ID_������� = 11
    ID_���Ƽ�� = 12
    ID_Ƕ��鿴 = 13
    ID_�Զ����� = 14
    
    ID_ʱ�䷶Χ = 2
    ID_һ�� = 21
    ID_���� = 22
    ID_���� = 23
    ID_���� = 24
    ID_һ�� = 25
    ID_���� = 26
    ID_���� = 27
    ID_���� = 28
    ID_�Զ��� = 29
    
End Enum

Private Const M_STR_COLNAME = "���;ҽ��ID;����;����;���;��Ŀ;��λ;������;��ǰ����;���ʱ��;ҽ������;�������"
'Private Const M_STR_COLNAME = "���,300,1;ҽ��ID,300,2;����,300,3;����,300,4;���,300,5;��Ŀ,300,,6;��λ,300,,7;������,300,8" _
'                                                  & ";��ǰ����,300,9;���ʱ��,300,10;ҽ������,300,11;�������,300,12"

Private Const M_STR_CFG = "[��ʷ]"

Private Type TPListCfg
    strSort As String
    strList As String
End Type

Public Event OnListLostFocus()
Public Event OnLoadCfg(ByRef strListCfg As String)
Public Event OnSaveCfg(ByVal strListCfg As String)
Public Event OnListMove()
Public Event OnListMouseClick(ByVal LngAdvice As Long, ByVal X As Long, ByVal Y As Long, ByVal blnClear As Boolean)
Public Event OnSelectStudy(ByVal LngAdvice As Long, ByVal strAdvice As String, ByVal blnEmbed As Boolean)
Public Event OnDoWork(ByVal LngAdvice As Long, ByVal strFuncName As String)
Public Event OnViewReport(ByVal LngAdvice As Long)
Public Event OnRefresh(ByVal lngCount As Long)

Property Let ListRow(value As Long)
    mlngRow = value
End Property

Public Function RefreshHistoryList(ByVal LngAdvice As Long, ByVal lngModule As Long, ByVal lngPatId As Long, ByVal blnDocPatient As Boolean, _
                            ByVal lngCur����ID As Long, ByVal strCanUse����IDs As String, ByVal lngLinkId As Long, _
                            ByVal blnAllDepts As Boolean, ByVal blnRelatingPatient As Boolean, Optional blnForce As Boolean = False, Optional lngNum As Long = 0) As Boolean
'blnDocPatient:�Ƿ�סԺ����
'lngNum:Ӥ�����

    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    Dim objItem As ListItem
    Dim iCount As Integer
    Dim strTime As String
    Dim dtBegin As Date
    Dim dtEnd As Date
    Dim blnMoved As Boolean
    Dim objControl As CommandBarControl
    Dim blnNoTime As Boolean
    Dim strValue As String

    On Error GoTo errHandle
    
    
    
    If lngModule = G_LNG_PATHSTATION_MODULE Then
        vsfStudy.TextMatrix(0, vsfStudy.ColIndex("����")) = "�����"
    End If
    
    If lngModule = G_LNG_VIDEOSTATION_MODULE Or lngModule = G_LNG_PATHSTATION_MODULE Then
        Set objControl = cbrMain.FindControl(, conMenu_Img_Look)
        objControl.Caption = "ͼ��"
        objControl.ToolTipText = "ͼ��"
        objControl.IconId = 10
        
        Set objControl = cbrMain.FindControl(, conMenu_Img_Contrast)
        objControl.IconId = 11
    End If
    
    If LngAdvice <= 0 Then Exit Function
    If mLngAdvice = LngAdvice And Not blnForce Then Exit Function

    mlngPatId = lngPatId
    mlngCur����ID = lngCur����ID
    mstrCanUse����IDs = strCanUse����IDs
    mlngLinkID = lngLinkId
    mblnAllDepts = blnAllDepts
    mblnRelatingPatient = blnRelatingPatient
    mlngModule = lngModule
    mLngAdvice = LngAdvice
    mblnDocPatient = blnDocPatient
    mlngBabyNum = lngNum
    
    Call SetVisible(ID_�������, mblnDocPatient, True)
    
    If blnAllDepts Then
        CheckCmd ID_���Ƽ��, True
        SetVisible ID_���Ƽ��, True, False
    Else
        SetVisible ID_���Ƽ��, True, True
    End If
    
    mlngRow = 0
    vsfStudy.Rows = 1
    blnNoTime = False
    
    '��ȡʱ�䷶Χ
    Select Case GetTime
        Case "һ��"
            dtBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")) - 30
            dtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
        Case "����"
            dtBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")) - 60
            dtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
        Case "����"
            dtBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")) - 90
            dtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
        Case "����"
            dtBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")) - 180
            dtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
        Case "һ��"
            dtBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")) - 365
            dtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
        Case "����"
            dtBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")) - 730
            dtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
        Case "����"
            dtBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")) - 1095
            dtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
        Case "����"
            blnNoTime = True
        Case "�Զ���"
            dtBegin = mdtBegin
            dtEnd = mdtEnd
    End Select
    
    '����ʱ��ʱ����dtBegin + 1��1899/12/31��ȥ�ж��Ƿ�ת�棬������dtBegin��ֱ��blnMoved = true
    If blnNoTime Then
        blnMoved = MovedByDate(dtBegin + 1)
    Else
        blnMoved = MovedByDate(dtBegin)
    End If
    
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        strSQL = "Select A.ID ҽ��ID,A.����ʱ��  ����ʱ��,A.ҽ������, B.ִ�й���, B.�������,C.����, C.Ӱ�����,C.�������,C.����,C.�������� ���ʱ��,E.����,E.�걾��λ " & _
               " From ����ҽ����¼ A,����ҽ������ B,Ӱ�����¼ C,������Ϣ D,������ĿĿ¼ E" & _
               " Where A.����id = [1] And A.���id Is Null And B.ҽ��ID=A.ID and a.����id = d.����id  " & _
               " AND A.ID=C.ҽ��ID(+) AND A.������ĿID = E.ID AND b.ִ�й��� >= 2 "
    Else
        strSQL = "Select A.ID ҽ��ID,A.����ʱ��  ����ʱ��,A.ҽ������, B.ִ�й���, B.�������,C.����� ����,E.����,E.�걾��λ,F.�������,F.Ӱ�����,F.����,F.�������� ���ʱ�� " & _
               " From ����ҽ����¼ A,����ҽ������ B,Ӱ�����¼ F,��������Ϣ C,������Ϣ D,������ĿĿ¼ E" & _
               " Where A.����id = [1] And A.���id Is Null And B.ҽ��ID=A.ID and a.����id = d.����id " & _
               " AND A.ID=C.ҽ��ID(+) AND A.������ĿID = E.ID and a.id=F.ҽ��ID(+) AND b.ִ�й��� >= 2 "
    End If
    
    If Not blnNoTime Then
        strSQL = strSQL & " AND B.����ʱ�� between [6] and [7]"
    End If
    
    '���μ��
    If IsCheck(ID_�������) And mblnDocPatient Then
        strSQL = strSQL & " And (A.������Դ=2 And A.��ҳID=D.��ҳID)"
    End If
    
    '���Ƽ��
    If blnAllDepts = False Then
        If Not IsCheck(ID_���Ƽ��) Then
            strSQL = strSQL & " And A.ִ�п���id+0 =[2] "
        Else
            strSQL = strSQL & " And  (A.ִ�п���id+0 <>[2] and B.ִ�й��� >= 5 or A.ִ�п���id+0 =[2]) "
        End If
    Else
        strSQL = strSQL & " And (Instr( [3],',' || A.ִ�п���id || ',' ) >0)"
    End If
    
    'Ӥ��
    strSQL = strSQL & " And NVL(A.Ӥ��,0) = [8]"
    
    '���ù������ˣ��Ų�ѯ����ID
    If blnRelatingPatient And lngLinkId <> 0 Then
        If mlngModule <> G_LNG_PATHOLSYS_NUM Then
            strSQL = strSQL & " union select A.ID ҽ��ID,A.����ʱ��  ����ʱ��,A.ҽ������, B.ִ�й���, B.�������,C.����, C.Ӱ�����,C.�������,C.����,C.�������� ���ʱ��,E.����,E.�걾��λ " & _
                " From ����ҽ����¼ A ,����ҽ������ B,Ӱ�����¼ C,������Ϣ D,������ĿĿ¼ E" & _
                " Where B.ҽ��ID=A.ID AND A.ID=C.ҽ��ID(+) and a.����id = d.����id AND A.������ĿID = E.ID AND b.ִ�й��� >= 2 AND A.id in (Select ҽ��ID from Ӱ�����¼ Where ����ID =[4]) "
        Else
            strSQL = strSQL & " union select A.ID ҽ��ID,A.����ʱ��  ����ʱ��,A.ҽ������, B.ִ�й���, B.�������,C.����� ����,E.����,E.�걾��λ,F.�������,F.Ӱ�����,F.����,F.�������� ���ʱ��" & _
                " From ����ҽ����¼ A,����ҽ������ B,Ӱ�����¼ F,��������Ϣ C,������Ϣ D,������ĿĿ¼ E" & _
                " Where A.id in (Select ҽ��ID from Ӱ�����¼ Where ����ID =[4]) And B.ҽ��ID=A.ID and a.id=C.ҽ��ID(+) and a.����id = d.����id AND A.������ĿID = E.ID and a.id=F.ҽ��ID(+) and b.ִ�й��� >= 2 "
        End If
        
        If Not blnNoTime Then
            strSQL = strSQL & " AND B.����ʱ�� between [6] and [7]"
        End If
        
        '���μ��
        If IsCheck(ID_�������) And mblnDocPatient Then
            strSQL = strSQL & " And (A.������Դ=2 And A.��ҳID=D.��ҳID)"
        End If
        
'        '���Ƽ��
'        If chkOtherDeptReport.Value <> 1 Then
'            strSql = strSql & " And c.ִ�п���id+0 in(select  ����id  from ������Ա where ��Աid = [5] union all select to_Number([2]) from dual) "
'        End If
        '���Ƽ��
        If blnAllDepts = False Then
            If Not IsCheck(ID_���Ƽ��) Then
                strSQL = strSQL & " And A.ִ�п���id+0 =[2] "
            Else
                strSQL = strSQL & " And  (A.ִ�п���id+0 <>[2] and B.ִ�й��� >= 5 or A.ִ�п���id+0 =[2]) "
            End If
        Else
            strSQL = strSQL & " And (Instr( [3],',' || A.ִ�п���id || ',' ) >0)"
        End If
        
        strSQL = strSQL & " And NVL(A.Ӥ��,0) = [8]"
    End If
    
    If blnMoved Then
        strTemp = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strTemp = Replace(strTemp, "����ҽ������", "H����ҽ������")
        strTemp = Replace(strTemp, "Ӱ�����¼", "HӰ�����¼")
        strTemp = Replace(strTemp, "���˼����Ϣ", "H���˼����Ϣ")
        strSQL = strSQL & vbNewLine & " Union ALL " & vbNewLine & strTemp
    End If
    strSQL = "Select * From (" & vbNewLine & strSQL & vbNewLine & ") Order By ����ʱ�� Asc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "", lngPatId, _
            lngCur����ID, "," & strCanUse����IDs & ",", lngLinkId, UserInfo.ID, dtBegin, dtEnd, lngNum)
    
    If rsTemp.RecordCount > 0 Then

        rsTemp.Filter = "ҽ��id <> " & LngAdvice
        
        With vsfStudy
            If mlngModule = G_LNG_PATHOLSYS_NUM Then .ColHidden(.ColIndex("���")) = True
                Do While Not rsTemp.EOF
                    .Rows = .Rows + 1
                    iCount = iCount + 1
                    
                    .TextMatrix(iCount, .ColIndex("ҽ��ID")) = Val(nvl(rsTemp!ҽ��ID))
        
                    .TextMatrix(iCount, .ColIndex("���")) = iCount
                    .TextMatrix(iCount, .ColIndex("����")) = nvl(rsTemp!����)
                    .TextMatrix(iCount, .ColIndex("����")) = nvl(rsTemp!����)
                    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
                        .TextMatrix(iCount, .ColIndex("���")) = nvl(rsTemp!Ӱ�����)
                    End If
                    .TextMatrix(iCount, .ColIndex("��Ŀ")) = nvl(rsTemp!����)
                    .TextMatrix(iCount, .ColIndex("��ǰ����")) = Decode(Val(nvl(rsTemp!ִ�й���, 0)), -1, "�Ѳ���", 0, "�ѵǼ�", 1, "�ѵǼ�", _
                                                        2, "�ѱ���", 3, "�Ѽ��", 4, "�ѱ���", 5, "�����", "�����")
                    .Cell(flexcpData, iCount, .ColIndex("��ǰ����")) = Val(nvl(rsTemp!ִ�й���, 0))
                    .TextMatrix(iCount, .ColIndex("������")) = IIf(Val(nvl(rsTemp!�������)) = 1, "��", "")
                    
                    
                    strTime = Format(rsTemp!���ʱ��, "yyyy-MM-dd hh:mm")
                    .TextMatrix(iCount, .ColIndex("�������")) = nvl(rsTemp!�������)
                    .TextMatrix(iCount, .ColIndex("���ʱ��")) = strTime
                    
                    
                    If UBound(Split(nvl(rsTemp!ҽ������), ":")) > 0 Then
                        .TextMatrix(iCount, .ColIndex("ҽ������")) = Split(nvl(rsTemp!ҽ������), ":")(0)
                        .TextMatrix(iCount, .ColIndex("��λ")) = Split(nvl(rsTemp!ҽ������), ":")(1)
                    Else
                        .TextMatrix(iCount, .ColIndex("ҽ������")) = nvl(rsTemp!ҽ������)
                        .TextMatrix(iCount, .ColIndex("��λ")) = ""
                    End If
                    
                    rsTemp.MoveNext
'                   If .Rows > 1 Then .Row = 1
                Loop
        End With
        
        If IsCheck(ID_�Զ�����) Then
            vsfStudy.WordWrap = True
            vsfStudy.AutoSize 1, vsfStudy.Cols - 1
        Else
            vsfStudy.WordWrap = False
            vsfStudy.AutoSize 1, vsfStudy.Cols - 1
        End If
    
    End If
    
    If Not mblnListCfgOk Then
        RaiseEvent OnLoadCfg(strValue)
        mblnListCfgOk = True
        
        If InStr(strValue, ";") > 0 Then
            mTPListCfg.strList = Split(strValue, ";")(1)
            mTPListCfg.strSort = Split(strValue, ";")(0)
            mTPListCfg.strSort = Replace(mTPListCfg.strSort, "[��ʷ]", "")
            Call DoLoadListCfg(mTPListCfg.strList)
            Call DoLoadListSort(mTPListCfg.strSort)
        Else
            mTPListCfg.strList = ""
            mTPListCfg.strSort = ""
        End If
    Else
        If mTPListCfg.strList <> "" Then
            Call DoLoadListCfg(mTPListCfg.strList)
        End If
        
        If mTPListCfg.strSort <> "" Then
            Call DoLoadListSort(mTPListCfg.strSort)
        End If
    End If
    
    
        
    RaiseEvent OnRefresh(rsTemp.RecordCount)
    RefreshHistoryList = True
    Exit Function
errHandle:
    MsgBox err.Description, vbOKOnly, gstrSysName
    err.Clear
End Function


Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnResult As Boolean
    
    On Error GoTo errHandle
    
    Select Case Control.ID
        Case conMenu_Img_Look, conMenu_Img_Contrast, conMenu_PacsReport_Open
            RaiseEvent OnDoWork(Val(vsfStudy.TextMatrix(vsfStudy.Row, vsfStudy.ColIndex("ҽ��ID"))), Control.Category)
        Case ID_�Զ���
            If mblnCustom Then Exit Sub
            CheckCmd Control.ID, Not Val(Control.Category) = 1
            mblnCustom = True
            blnResult = frmSetTime.ShowSetTime(mdtBegin, mdtEnd, Me)
            mblnCustom = False

            Call RefreshHistoryList(mLngAdvice, mlngModule, mlngPatId, mblnDocPatient, mlngCur����ID, mstrCanUse����IDs, mlngLinkID, mblnAllDepts, mblnRelatingPatient, True, mlngBabyNum)
        Case Else
            CheckCmd Control.ID, Not Val(Control.Category) = 1
            Call RefreshHistoryList(mLngAdvice, mlngModule, mlngPatId, mblnDocPatient, mlngCur����ID, mstrCanUse����IDs, mlngLinkID, mblnAllDepts, mblnRelatingPatient, True, mlngBabyNum)
    End Select
    Exit Sub
errHandle:
    MsgBox err.Description, vbOKOnly, gstrSysName
    err.Clear
End Sub

Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error Resume Next
    
    Select Case Control.ID
        Case conMenu_Img_Look, conMenu_Img_Contrast, conMenu_PacsReport_Open
            If vsfStudy.ColIndex("ҽ��ID") = -1 Then Exit Sub
            Control.Enabled = Val(vsfStudy.TextMatrix(vsfStudy.Row, vsfStudy.ColIndex("ҽ��ID"))) > 0
            Control.Visible = Not IsCheck(ID_Ƕ��鿴)
            If Control.Visible And Control.Enabled Then
                If Control.ID = conMenu_PacsReport_Open Then
                    Control.Enabled = mblnReportEnable
                Else
                    Control.Enabled = mblnImageEnable
                End If
            End If
    End Select
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    
    mblnImageEnable = False
    mblnReportEnable = False
    
    mblnListCfgOk = False
    Call InitCommandBars
    Call GridInit(M_STR_COLNAME)
    Call SetFontSize(gbytFontSize)
    
    
    CheckCmd ID_���Ƽ��, Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���Ƽ��", "0")) = 1
    CheckCmd ID_�������, Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�������", "0")) = 1
    CheckCmd ID_Ƕ��鿴, Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "Ƕ��鿴", "0")) = 1
    CheckCmd ID_�Զ�����, Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�Զ�����", "0")) = 1
    
    mdtBegin = CDate(Format(zlDatabase.Currentdate - 365, "yyyy-mm-dd 00:00:00"))
    mdtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
    Exit Sub
errHandle:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    
    cbrMain.RecalcLayout
    cbrMain.GetClientRect lngLeft, lngTop, lngRight, lngBottom
    
    vsfStudy.Left = 0
    vsfStudy.Top = lngTop
    vsfStudy.Width = Me.ScaleWidth
    vsfStudy.Height = Me.ScaleHeight - vsfStudy.Top
End Sub

Private Sub GridInit(strColName As String)
On Error GoTo errH
    '��ʼ�������б�
    Dim i As Integer
    Dim lngCount As Long
    Dim arrData() As String
    
    arrData = Split(strColName, ";")
    lngCount = UBound(arrData) + 1
    
    With vsfStudy
    
        .Cols = lngCount
        .FixedRows = 1
        .FixedCols = 0
        .RowHeightMin = 320
'        .Cell(flexcpAlignment, 0, 0, 0, lngCount - 1) = flexAlignCenterCenter

        '���һ���Զ�������б�
        .AllowUserResizing = flexResizeColumns
        .ExtendLastCol = True
        .AutoResize = True
        .ExplorerBar = 7 '������ͷ�϶�������
        .AutoSizeMode = flexAutoSizeRowHeight
    
        .WordWrap = True
        .AutoSizeMouse = True
        .SelectionMode = flexSelectionByRow
        .ScrollTrack = True
        
        For i = 0 To lngCount - 1
            .TextMatrix(0, i) = arrData(i)
            .ColKey(i) = arrData(i)
        Next
        
        .Rows = 1
        If .Rows > 1 Then .RowSel = 1
        
        .ColHidden(.ColIndex("ҽ��ID")) = True '����ҽ��ID
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Public Sub SetFontSize(ByVal bytFontSize As Byte)
    Dim CtlFont As StdFont
    
    Set CtlFont = New StdFont
    CtlFont.Size = bytFontSize
    
    Call SetColWithd(bytFontSize)
    
    vsfStudy.FontSize = bytFontSize
    Set cbrMain.Options.Font = CtlFont
    
    If bytFontSize = 9 Then
        cbrMain.Options.SetIconSize True, 16, 16
    ElseIf bytFontSize = 12 Then
        cbrMain.Options.SetIconSize True, 20, 20
    ElseIf bytFontSize = 15 Then
        cbrMain.Options.SetIconSize True, 24, 24
    End If
    Call Form_Resize
End Sub

Private Sub SetColWithd(ByVal bytSize As Long)

    If mblnListCfgOk Then Exit Sub
    With vsfStudy
        Select Case bytSize
            Case 9
                .ColWidth(.ColIndex("���")) = 500
                .ColWidth(.ColIndex("��λ")) = 1000
                .ColWidth(.ColIndex("��ǰ����")) = 900
                .ColWidth(.ColIndex("����")) = 700
                .ColWidth(.ColIndex("���")) = 700
                .ColWidth(.ColIndex("����")) = 700
                .ColWidth(.ColIndex("���ʱ��")) = 1600
                .ColWidth(.ColIndex("��Ŀ")) = 1200
                .ColWidth(.ColIndex("������")) = 800
            Case 12
                .ColWidth(.ColIndex("���")) = 600
                .ColWidth(.ColIndex("��λ")) = 1250
                .ColWidth(.ColIndex("��ǰ����")) = 1100
                .ColWidth(.ColIndex("����")) = 900
                .ColWidth(.ColIndex("���")) = 900
                .ColWidth(.ColIndex("����")) = 900
                .ColWidth(.ColIndex("���ʱ��")) = 2200
                .ColWidth(.ColIndex("��Ŀ")) = 1450
                .ColWidth(.ColIndex("������")) = 1000
            Case 15
                .ColWidth(.ColIndex("���")) = 700
                .ColWidth(.ColIndex("��λ")) = 1500
                .ColWidth(.ColIndex("��ǰ����")) = 1300
                .ColWidth(.ColIndex("����")) = 1100
                .ColWidth(.ColIndex("���")) = 1100
                .ColWidth(.ColIndex("����")) = 1100
                .ColWidth(.ColIndex("���ʱ��")) = 2800
                .ColWidth(.ColIndex("��Ŀ")) = 1700
                .ColWidth(.ColIndex("������")) = 1200
        End Select
    End With
End Sub

Private Sub vsfStudy_AfterMoveColumn(ByVal Col As Long, Position As Long)
    mTPListCfg.strList = GetListHeadString
    RaiseEvent OnSaveCfg(M_STR_CFG & mTPListCfg.strSort & ";" & mTPListCfg.strList)
End Sub

Private Sub vsfStudy_AfterSort(ByVal Col As Long, Order As Integer)
    Dim strName As String
    Dim i As Integer
    
    For i = 1 To vsfStudy.Rows - 1
        vsfStudy.TextMatrix(i, vsfStudy.ColIndex("���")) = i
    Next
    
    strName = vsfStudy.TextMatrix(0, Col)
    mTPListCfg.strSort = strName & "," & Order
    
    RaiseEvent OnSaveCfg(M_STR_CFG & mTPListCfg.strSort & ";" & mTPListCfg.strList)
End Sub

Private Sub vsfStudy_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    mTPListCfg.strList = GetListHeadString
    RaiseEvent OnSaveCfg(M_STR_CFG & mTPListCfg.strSort & ";" & mTPListCfg.strList)
End Sub

Private Sub vsfStudy_Click()
    Call vsfStudy_RowColChange
End Sub

Private Sub vsfStudy_DblClick()
    Dim lngRow As Long
    Dim strAdvice As String
    Dim intCol As Integer
    
    On Error GoTo errHandle
    
    lngRow = vsfStudy.MouseRow
    intCol = vsfStudy.ColIndex("ҽ��ID")

    If lngRow = 0 Or intCol = -1 Then Exit Sub
    
    If IsCheck(ID_Ƕ��鿴) Then Exit Sub
    
    If lngRow <= 0 Then Exit Sub
    If Val(vsfStudy.TextMatrix(lngRow, intCol)) <= 0 Then Exit Sub
    
    RaiseEvent OnViewReport(Val(vsfStudy.TextMatrix(lngRow, intCol)))
    
    Exit Sub
errHandle:
    MsgBox err.Description, vbExclamation, "��ʾ"
    err.Clear
End Sub

Private Sub vsfStudy_LostFocus()
    RaiseEvent OnListLostFocus
End Sub

Private Sub vsfStudy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OnListMove
End Sub

Private Sub vsfStudy_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errH
    Dim Popup As CommandBar
    Dim Control As CommandBarControl
    Dim intCol As Integer
    
    Dim lngID As Long
    Dim lngRow As Long
    
    
    If Button = 2 Then
        If IsCheck(ID_Ƕ��鿴) Then Exit Sub
        Set Popup = cbrMain.Add("�Ҽ��˵�", xtpBarPopup)
        With Popup.Controls
            Set Control = .Add(xtpControlButton, conMenu_Img_Look, "��Ƭ"): Control.IconId = 5: Control.Category = "��Ƭ"
            Set Control = .Add(xtpControlButton, conMenu_Img_Contrast, "�Ա�"): Control.IconId = 6: Control.Category = "�Ա�"
            Set Control = .Add(xtpControlButton, conMenu_PacsReport_Open, "�鿴����"): Control.IconId = 7: Control.Category = "�鿴����"
        End With
        
        Call Popup.ShowPopup
        
    ElseIf Button = 1 Then
        intCol = vsfStudy.ColIndex("ҽ��ID")
    
        If intCol = -1 Then Exit Sub
    
        lngRow = vsfStudy.MouseRow
        If lngRow > 0 Then
            lngID = Val(vsfStudy.TextMatrix(lngRow, intCol))
        End If
        
        If vsfStudy.MouseRow > 0 Then
            RaiseEvent OnListMouseClick(lngID, X, Y, False)
        Else
            RaiseEvent OnListMouseClick(0, X, Y, True)
        End If
    End If
errH:
End Sub

Private Sub vsfStudy_RowColChange()
    Dim lngRow As Long
    Dim i As Long
    Dim intCol As Integer
    Dim strAdvice As String
    
    On Error GoTo errHandle
    
    lngRow = vsfStudy.Row
    intCol = vsfStudy.ColIndex("ҽ��ID")
    
    If lngRow <= 0 Or intCol = -1 Then Exit Sub
    
    If mlngRow = lngRow Then Exit Sub
    If Val(vsfStudy.TextMatrix(lngRow, intCol)) <= 0 Then Exit Sub
    mlngRow = lngRow
    
    For i = 1 To vsfStudy.Rows - 1
        If Val(vsfStudy.TextMatrix(i, intCol)) > 0 Then
            strAdvice = strAdvice & IIf(Len(strAdvice) = 0, "", "|") & vsfStudy.TextMatrix(i, intCol)
        End If
    Next
    
    RaiseEvent OnSelectStudy(Val(vsfStudy.TextMatrix(lngRow, intCol)), strAdvice, IsCheck(ID_Ƕ��鿴))
    
    Exit Sub
errHandle:
    MsgBox err.Description, vbExclamation, "��ʾ"
    err.Clear
End Sub

Public Sub ClearData()
'�������
    vsfStudy.Rows = 1
    
    mLngAdvice = 0
End Sub

Private Sub InitCommandBars()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl, cbrPopControl As CommandBarControl
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    Set Me.cbrMain.Icons = imgList.Icons
    
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True  '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .UseSharedImageList = False
    End With
    
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.Visible = False
    
    '����������
    Set cbrToolBar = Me.cbrMain.Add("������", xtpBarTop)
    cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = True
    cbrToolBar.Closeable = False
    cbrToolBar.ContextMenuPresent = False
    
    With cbrToolBar.Controls
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Look, "��Ƭ"): cbrControl.IconId = 5: cbrControl.Category = "��Ƭ": cbrControl.ToolTipText = "��Ƭ"
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Contrast, "�Ա�"): cbrControl.IconId = 6: cbrControl.Category = "�Ա�": cbrControl.ToolTipText = "�Ա�"
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_Open, "����"): cbrControl.IconId = 7: cbrControl.Category = "�鿴����": cbrControl.ToolTipText = "�鿴����"
    
        'ʱ��.........................................................
        Set cbrControl = .Add(xtpControlButtonPopup, ID_ʱ�䷶Χ, "����"): cbrControl.IconId = 4: cbrControl.ToolTipText = "����"
        cbrControl.BeginGroup = True

        Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, ID_һ��, "һ��"): objControl.IconId = 8
        Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, ID_����, "����"): objControl.IconId = 8
        Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, ID_����, "����"): objControl.IconId = 8
        Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, ID_����, "����"): objControl.IconId = 8
        Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, ID_һ��, "һ��"): objControl.IconId = 9: objControl.Category = 1
        Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, ID_����, "����"): objControl.IconId = 8
        Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, ID_����, "����"): objControl.IconId = 8
        Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, ID_����, "����"): objControl.IconId = 8
        Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, ID_�Զ���, "�Զ���"): objControl.IconId = 8
        
        For Each objControl In cbrControl.CommandBar.Controls
            objControl.CloseSubMenuOnClick = False
        Next
        
        Set cbrControl = .Add(xtpControlButtonPopup, ID_�������, "ѡ��"): cbrControl.IconId = 3: cbrControl.ToolTipText = "ѡ��"
        
        
        Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, ID_�������, "�������"): cbrPopControl.IconId = 1
        Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, ID_���Ƽ��, "���Ƽ��"): cbrPopControl.IconId = 1
        Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, ID_Ƕ��鿴, "Ƕ��鿴"): cbrPopControl.IconId = 1
        Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, ID_�Զ�����, "�Զ�����"): cbrPopControl.IconId = 1

        For Each cbrPopControl In cbrControl.CommandBar.Controls
            cbrPopControl.CloseSubMenuOnClick = False
        Next
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
End Sub

Private Function IsCheck(ByVal fltID As FilterID) As Boolean
    Dim objControl As CommandBarControl
    Dim obj As CommandBarControl
    
    If fltID > 10 And fltID < 20 Then
        Set objControl = cbrMain.FindControl(, ID_�������)
    ElseIf fltID > 20 And fltID < 30 Then
        Set objControl = cbrMain.FindControl(, ID_ʱ�䷶Χ)
    End If
    Set obj = objControl.CommandBar.FindControl(, fltID)
    IsCheck = Val(obj.Category) = 1
End Function


Private Sub CheckCmd(ByVal fltID As FilterID, ByVal blnCheck As Boolean)
    Dim objControl As CommandBarControl
    Dim obj As CommandBarControl
    Dim i As Long
    
    Select Case fltID
        Case ID_�������, ID_Ƕ��鿴, ID_���Ƽ��, ID_�Զ�����
            Set objControl = cbrMain.FindControl(, ID_�������)
            Set obj = objControl.CommandBar.FindControl(, fltID)
            obj.IconId = IIf(blnCheck, 2, 1)
            obj.Category = IIf(blnCheck, 1, 0)
        Case ID_����, ID_����, ID_����, ID_����, ID_����, ID_һ��, ID_һ��, ID_�Զ���, ID_����
            If blnCheck Then
                Set objControl = cbrMain.FindControl(, ID_ʱ�䷶Χ)
                For i = 21 To 29
                    
                    Set obj = objControl.CommandBar.FindControl(, i)
                    If fltID <> i Then
                        obj.IconId = 8
                        obj.Category = 0
                    Else
                        obj.IconId = 9
                        obj.Category = 1
                    End If
                Next
            End If
    End Select
End Sub

Private Sub SetVisible(ByVal fltID As FilterID, ByVal blnVisible As Boolean, ByVal blnEnabled As Boolean)
    Dim objControl As CommandBarControl
    Dim obj As CommandBarControl
    
    If fltID > 10 And fltID < 20 Then
        Set objControl = cbrMain.FindControl(, ID_�������)
    ElseIf fltID > 20 And fltID < 30 Then
        Set objControl = cbrMain.FindControl(, ID_ʱ�䷶Χ)
    End If
    Set obj = objControl.CommandBar.FindControl(, fltID)
    obj.Visible = blnVisible
    obj.Enabled = blnEnabled
End Sub

Private Function GetTime() As String
    Dim objControl As CommandBarControl
    Dim obj As CommandBarControl
    Dim i As Long
    
    Set objControl = cbrMain.FindControl(, ID_ʱ�䷶Χ)
    
    For i = 21 To 29
        Set obj = objControl.CommandBar.FindControl(, i)
        If Val(obj.Category) = 1 Then
            GetTime = obj.Caption
            Exit Function
        End If
    Next
    
    GetTime = "һ��"
End Function

Private Function IsImageEnable(ByVal LngAdvice As Long) As Boolean
On Error GoTo errH
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "select ���UID from Ӱ�����¼ where  ҽ��id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���UID", LngAdvice)
    
    If rsTemp.RecordCount <= 0 Then Exit Function
    
    IsImageEnable = Len(nvl(rsTemp!���UID)) > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function IsReportEnable(ByVal LngAdvice As Long) As Boolean
On Error GoTo errH
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    strSQL = "select count(1) ���� from ����ҽ������ where  ҽ��id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ����", LngAdvice)
    
    If rsTemp.RecordCount <= 0 Then Exit Function
    
    IsReportEnable = Val(nvl(rsTemp!����)) > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub Free()
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���Ƽ��", IIf(IsCheck(ID_���Ƽ��), 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�������", IIf(IsCheck(ID_�������), 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "Ƕ��鿴", IIf(IsCheck(ID_Ƕ��鿴), 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�Զ�����", IIf(IsCheck(ID_�Զ�����), 1, 0)
End Sub

Private Sub vsfStudy_SelChange()
On Error GoTo errHandle
    Dim lngAdviceID As Long
    Dim intCol As Integer
    
    
    intCol = vsfStudy.ColIndex("ҽ��ID")
    If intCol = -1 Then Exit Sub
    
    lngAdviceID = Val(vsfStudy.TextMatrix(vsfStudy.Row, intCol))
    
    mblnReportEnable = IsReportEnable(lngAdviceID)
    mblnImageEnable = IsImageEnable(lngAdviceID)
Exit Sub
errHandle:
    MsgBox err.Description, vbOKOnly, gstrSysName
End Sub

Private Function GetListHeadString() As String
'�õ���������: ����,���,�Ƿ���ʾ  ����  "���,1000,1|ִ�й���,2000,0|"
On Error GoTo errH
    Dim i As Integer
    Dim strTemp As String
    Dim strName As String
    Dim lngWidth As Long
    Dim blnShow As Boolean
    
    For i = 0 To vsfStudy.Cols - 1
        
        strName = vsfStudy.TextMatrix(0, i)
        lngWidth = vsfStudy.ColWidth(i)
        blnShow = Not vsfStudy.ColHidden(i)
        
        If Len(strTemp) > 0 Then
            strTemp = strTemp & "|"
        End If
        
        strTemp = strTemp & strName & "," & lngWidth & "," & blnShow
    Next

    GetListHeadString = strTemp
    
    Exit Function
errH:
    err.Raise -1, "��ʷ���", "[��ȡ��ͷ����]" & vbCrLf & err.Description
    Resume
End Function

Private Sub DoLoadListCfg(ByVal strcfg As String)
'�ָ���˳��Ϳ��
On Error GoTo errH
    Dim i As Integer, j As Integer
    Dim strName As String
    Dim lngW As Long
    Dim strCol() As String
    Dim intubound As Integer
    Dim blnHide As Boolean
    
    strCol = Split(strcfg, "|")
    intubound = UBound(strCol)
    
    With vsfStudy
        For i = 0 To intubound - 1
            strName = Split(strCol(i), ",")(0)
            lngW = Split(strCol(i), ",")(1)
            blnHide = Not Split(strCol(i), ",")(2)

            If strName <> .TextMatrix(0, i + 1) Then
                For j = 0 To .Cols - 1
                    If strName = .TextMatrix(0, j) Then
                        .ColPosition(j) = i
                        .ColWidth(i) = lngW
                        .ColHidden(i) = blnHide
                        Exit For
                    End If
                Next
            Else
                .ColWidth(i) = lngW
                .ColHidden(i) = blnHide
            End If

        Next
    End With
    
    Exit Sub
errH:
    err.Raise -1, "�б���Ի�����", "[DoLoadListCfg]" & vbCrLf & err.Description
    Resume
End Sub

Private Sub DoLoadListSort(ByVal strcfg As String)
'�ָ�����
On Error GoTo errH
    Dim strName As String
    Dim intWay As Integer
    Dim intPos As Integer
    Dim intCol As Integer
    Dim i As Integer
    
    intPos = InStr(strcfg, ",")
    If intPos = 0 Then Exit Sub
    
    strName = Split(strcfg, ",")(0)
    intWay = Val(Split(strcfg, ",")(1))
    
    With vsfStudy
        For i = 1 To .Cols - 1
            If strName = .TextMatrix(0, i) Then
                intCol = i
                Exit For
            End If
        Next
         
        .Col = intCol
        .Sort = intWay
        
        For i = 1 To .Rows - 1
            .TextMatrix(i, vsfStudy.ColIndex("���")) = i
        Next
    End With
    
    Exit Sub
errH:
    err.Raise -1, "�б���Ի�����", "[DoLoadListSort]" & vbCrLf & err.Description
    Resume
End Sub



