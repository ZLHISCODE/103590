VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDiseaseReportMan 
   Caption         =   "�����걨����"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10230
   Icon            =   "frmDiseaseReportMan.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7440
   ScaleWidth      =   10230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   3330
      Left            =   120
      TabIndex        =   1
      Top             =   630
      Width           =   6660
      _Version        =   589884
      _ExtentX        =   11747
      _ExtentY        =   5874
      _StockProps     =   0
      BorderStyle     =   2
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7065
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDiseaseReportMan.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15161
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgTemp 
      Height          =   900
      Left            =   240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5340
      Visible         =   0   'False
      Width           =   1080
      _cx             =   1905
      _cy             =   1587
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
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
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgInfo 
      Height          =   6300
      Left            =   7020
      TabIndex        =   3
      Top             =   630
      Width           =   3135
      _cx             =   5530
      _cy             =   11112
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483637
      GridColorFixed  =   -2147483638
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   1
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   1545
      Top             =   5355
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseReportMan.frx":0E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseReportMan.frx":11B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseReportMan.frx":1550
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseReportMan.frx":18EA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   285
      Top             =   75
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmDiseaseReportMan.frx":1C84
      Left            =   960
      Top             =   210
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmDiseaseReportMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------
'����
'-----------------------------------------------------
Private Enum mCol
    ͼ�� = 0: ID: ״̬: ����: ����: �����: ����: �Ա�: ����: �ʱ��: ���: ��Ϣ: ����ת��: ����ID: ��ҳID: �ļ�ID: �༭��ʽ
End Enum
Const conPane_Reports = 1
Const conPane_Preview = 2
Const conPane_AppInfo = 3

Private mobjDoc As cEPRDocument
Private mobjRichEMR As Object
Private mobjInfection As Object
'-----------------------------------------------------
'�������
'-----------------------------------------------------
Private mstrPrivs As String             '��ǰʹ����Ȩ�޴�
Private mstrFiles As String             '��������ı����ļ�
Private mintDates As Integer            'Ĭ�ϲ鿴�����¼��������Ϊ0ʱ��˵���������ִ�в������ã�Ҫ�����ڷ�Χ�鿴
Private mstrDateFrom As String          '����Χ�鿴��ʼ���ڣ���mintDates=0ʱ��Ч
Private mstrDateTo As String            '����Χ�鿴��ֹ���ڣ���mintDates=0ʱ��Ч

Private mfrmPreview As frmDockEPRContent  '��������Ԥ������
Private mstrCurId As String               '��ǰ��¼ID EMR���ID���ַ���
Private mstrContent As String             '�²�����XML����
Private mintState As Integer            '��ǰ��¼״̬
Private mblnCurMoved As Boolean         '��ǰ��¼ת��״̬ 0-δת�� 1-��ת��
'-----------------------------------------------------

Private Function zlRefList(Optional strCurId As String, Optional strSender As String, Optional strPatient As String, Optional lngOutNo As Long, Optional lngInNo As Long) As Long
'���ܣ�ˢ��װ����������ļ������棬����λ��ָ���ļ�¼��
'������strCurId ��λID ���ա����ͣ�������ˢ��ʱ����
'       ���ָ����������������š�סԺ����ʹ��ʱ������,�����˵��ĸ����������0��1��
'       strSender  ������ ���Ҵ���
'       strPatient �������� ���Ҵ���
'       lngOutNo   �����   ���Ҵ���
'       lngInNo    סԺ��   ���Ҵ���
Dim blnMoved As Boolean, strTemp As String, strFiles As String, i As Integer, strReturn As String
Dim rsTemp As New ADODB.Recordset, rsData As New ADODB.Recordset
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow

    Err = 0: On Error GoTo errHand
    
    If Trim(mstrFiles) = "" Then Exit Function
    Me.rptList.Records.DeleteAll
    
    For i = 0 To UBound(Split(mstrFiles, ","))
        If IsNumeric(Split(mstrFiles, ",")(i)) Then
            strFiles = strFiles & "," & Split(mstrFiles, ",")(i)
        End If
    Next
    If strFiles <> "" Then
        strFiles = Mid(strFiles, 2)
    End If
    
    If strFiles <> "" Then
        If mintDates <> 0 Then '���ָ����������������š�סԺ����ʹ��ʱ������
            gstrSQL = "And l.���ʱ�� >= trunc(Sysdate - [1])"
        Else
            blnMoved = MovedByDate(CDate(mstrDateFrom))
            gstrSQL = "And l.���ʱ�� Between To_Date([2],'yyyy-mm-dd') And To_Date([3],'yyyy-mm-dd')+1-1/24/60/60"
        End If
        
        If strPatient <> "" Or lngOutNo <> 0 Or lngInNo <> 0 Then
            gstrSQL = "And l.���ʱ�� is not null"
        End If
        
        gstrSQL = "Select l.Id,l.�ļ�id,l.����ID,l.��ҳID, l.�������� As ����, Decode(l.������Դ, 1, '����: ', 2, 'סԺ: ', '') || d.���� As ����," & _
                "        Decode(l.������Դ, 1, p.�����, 2, p.סԺ��) As �����, Nvl(l.����,p.����) as ����, Nvl(l.�Ա�,p.�Ա�) as �Ա�, Nvl(l.����,p.����) as ����, " & _
                "        To_Char(l.���ʱ��, 'yyyy-mm-dd hh24:mi') As �ʱ��, l.������ As ���,l.�༭��ʽ," & _
                "        Decode(l.״̬, -1, Decode(Sign(l.����ʱ�� - l.�վ�ʱ��), 1, 0, -1), l.״̬) As ״̬," & _
                "        l.�վ��� || '|' || To_Char(l.�վ�ʱ��, 'yyyy-mm-dd hh24:mi') || '|' || l.�վ�˵�� || '|' || l.������ || '|' ||" & _
                "        To_Char(l.����ʱ��, 'yyyy-mm-dd hh24:mi') || '|' || l.���͵�λ || '|' || l.���ͱ�ע || '|' || l.�Ǽ��� || '|' ||" & _
                "        To_Char(l.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi') || '|' || l.ְҵ || '|' || l.��ͥ��ַ || '|' || l.��ͥ�绰 || '|' ||" & _
                "        To_Char(l.��������, 'yyyy-mm-dd') || '|' || To_Char(l.ȷ������,'yyyy-mm-dd') || '|' ||" & _
                "        l.�������1 || '|' || l.�������2 || '|' || l.���ע As ��Ϣ,0 as ����ת��" & _
                " From (Select l.Id,l.�ļ�id,l.����ID,l.��ҳID, l.��������, l.������Դ, l.����id, l.���ʱ��, l.������, l.����ʱ��,l.�༭��ʽ," & _
                "               Nvl(s.����״̬, 0) As ״̬, s.�վ���, s.�վ�ʱ��, s.�վ�˵��, s.������, s.����ʱ��, s.���͵�λ, s.���ͱ�ע," & _
                "               s.�Ǽ��� , s.�Ǽ�ʱ��, s.����, s.�Ա�, s.����, s.ְҵ, s.��ͥ��ַ, s.��ͥ�绰, s.��������, s.ȷ������, " & _
                "               s.�������1, s.�������2, s.���ע" & _
                "        From ���Ӳ�����¼ l, �����걨��¼ s" & _
                "        Where l.Id = s.�ļ�id(+) And l.�������� = 5 And l.�ļ�id In (" & strFiles & ") " & gstrSQL & _
                IIf(strSender = "", "", "And s.������=[4]") & _
                "       ) l,������Ϣ p, ���ű� d" & IIf(lngInNo = 0, "", ",(Select Distinct ����id,סԺ�� From ������ҳ Where סԺ�� = [7]) A ") & _
                " Where l.����id = p.����id And l.����id = d.Id" & IIf(strPatient = "", "", " And p.����=[5]") & _
                IIf(lngOutNo = 0, "", " And p.�����=[6]") & IIf(lngInNo = 0, "", " And a.סԺ��=[7] And a.����ID=p.����ID")
        If blnMoved Then
            strTemp = Replace(gstrSQL, "0 as ����ת��", "1 as ����ת��")
            strTemp = Replace(strTemp, "���Ӳ�����¼", "H���Ӳ�����¼")
            strTemp = Replace(strTemp, "�����걨��¼", "H�����걨��¼")
            gstrSQL = gstrSQL & " Union All " & strTemp
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mintDates, mstrDateFrom, mstrDateTo, strSender, strPatient, lngOutNo, lngInNo)
    
        Do While Not rsTemp.EOF
            Set rptRcd = Me.rptList.Records.Add()
            Set rptItem = rptRcd.AddItem(CStr(rsTemp!״̬))
            Select Case rptItem.Value
            Case 0: rptItem.Icon = 0
            Case 1: rptItem.Icon = 1
            Case -1: rptItem.Icon = 2
            Case 2: rptItem.Icon = 3
            End Select
            rptRcd.AddItem CStr(rsTemp!ID)
            Select Case rsTemp!״̬
            Case 0: rptRcd.AddItem CStr("a)����д�ļ�������")
            Case 1: rptRcd.AddItem CStr("b)�ѽ��յļ�������")
            Case -1: rptRcd.AddItem CStr("c)�Ѿ��յļ�������")
            Case 2: rptRcd.AddItem CStr("d)�ѱ��͵ļ�������")
            Case Else: rptRcd.AddItem ""
            End Select
            rptRcd.AddItem CStr(rsTemp!����)
            rptRcd.AddItem CStr(rsTemp!����)
            rptRcd.AddItem CStr("" & rsTemp!�����)
            rptRcd.AddItem CStr(rsTemp!����)
            rptRcd.AddItem CStr("" & rsTemp!�Ա�)
            rptRcd.AddItem CStr("" & rsTemp!����)
            rptRcd.AddItem CStr(NVL(rsTemp!�ʱ��))
            rptRcd.AddItem CStr(NVL(rsTemp!���))
            rptRcd.AddItem CStr(rsTemp!��Ϣ)
            rptRcd.AddItem CStr(rsTemp!����ת��)
            rptRcd.AddItem CStr(rsTemp!����ID)
            rptRcd.AddItem CStr(rsTemp!��ҳID)
            rptRcd.AddItem CStr(rsTemp!�ļ�ID)
            rptRcd.AddItem CStr(NVL(rsTemp!�༭��ʽ, 0))
            rsTemp.MoveNext
        Loop
    End If
    
    If Not gobjEmr Is Nothing Then
        strFiles = ""
        For i = 0 To UBound(Split(mstrFiles, ","))
            If Not IsNumeric(Split(mstrFiles, ",")(i)) Then
                strFiles = strFiles & ",Hextoraw('" & Split(mstrFiles, ",")(i) & "')"
            End If
        Next
        If strFiles <> "" Then
            strFiles = Mid(strFiles, 2)
        End If
        
        If strFiles <> "" Then
            If mintDates <> 0 Then
                gstrSQL = "l.complete_time >= trunc(Sysdate - :dates)"
            Else
                gstrSQL = "l.complete_time Between To_Date(:datef,'yyyy-mm-dd') And To_Date(:datet,'yyyy-mm-dd')+1-1/24/60/60"
            End If
            
            If strPatient <> "" Or lngOutNo <> 0 Or lngInNo <> 0 Then
                gstrSQL = "l.complete_time is not null"
            End If
            
            gstrSQL = "Select Rawtohex(m.Id) ID,Rawtohex(l.Antetype_Id) AntetypeId, m.Title ����, Decode(o.Title, '�������', 1, 2) ������Դ,l.Completor �����, m.editor �༭��, p.Code ����id," & vbNewLine & _
                        "To_Char(m.Edit_Time, 'yyyy-mm-dd hh24:mi:ss') ����ʱ��, To_Char(l.Complete_Time, 'yyyy-mm-dd hh24:mi:ss') ���ʱ��, To_Char(n.Begin_Time, 'yyyy-mm-dd hh24:mi:ss') �¼�ʱ��" & vbNewLine & _
                    "From Bz_Doc_Tasks L, Bz_Doc_Log M, Bz_Act_Log N, Action_List O, Bz_Master_Codes P" & vbNewLine & _
                    "Where " & gstrSQL & " And l.Antetype_Id In (" & strFiles & ") And l.Real_Doc_Id = m.Id And m.Status >= 2 And" & vbNewLine & _
                    "      m.Actlog_Id = n.Id And n.Action_Id = o.Id And n.Master_Id = p.Master_Id And p.Kind = '����ID'"
            If mintDates <> 0 Then
                strReturn = gobjEmr.OpenSQLRecordset(gstrSQL, IIf(strPatient <> "" Or lngOutNo <> 0 Or lngInNo <> 0, "", mintDates & "^11^dates"), rsTemp)
            Else
                strReturn = gobjEmr.OpenSQLRecordset(gstrSQL, IIf(strPatient <> "" Or lngOutNo <> 0 Or lngInNo <> 0, "", mstrDateFrom & "^16^datef|" & mstrDateTo & "^16^datet"), rsTemp)
            End If
            
            If strReturn = "" Then
            Do Until rsTemp.EOF
                gstrSQL = "Select '" & rsTemp!ID & "' As ID, p.����id, '" & rsTemp!AntetypeId & "' As �ļ�id, q.��ҳid," & vbNewLine & _
                            "       '" & rsTemp!���� & "' As ����, Decode(" & rsTemp!������Դ & ", 1, '����:', 2, 'סԺ:') || a.���� As ����, Decode(" & rsTemp!������Դ & ", 1, p.�����, 2, p.סԺ��) �����," & vbNewLine & _
                            "       '" & Format(rsTemp!���ʱ��, "yyyy-mm-dd HH:MM") & "' As �ʱ��, p.����, p.�Ա�, p.����,0 As ����ת��, '" & rsTemp!�༭�� & "' As ���,3 as �༭��ʽ," & vbNewLine & _
                            "Nvl((Select Decode(Nvl(s.����״̬, 0), -1," & vbNewLine & _
                            "                 Decode(Sign(To_Date('2014-12-22 10:47:09', 'yyyy-mm-dd hh24:mi:ss') - s.�վ�ʱ��), 1, 0, -1)," & vbNewLine & _
                            "                 Nvl(s.����״̬, 0)) || '|' || s.�վ��� || '|' || To_Char(s.�վ�ʱ��, 'yyyy-mm-dd hh24:mi') || '|' || s.�վ�˵�� || '|' ||" & vbNewLine & _
                            "          s.������ || '|' || To_Char(s.����ʱ��, 'yyyy-mm-dd hh24:mi') || '|' || s.���͵�λ || '|' || s.���ͱ�ע || '|' || s.�Ǽ��� || '|' ||" & vbNewLine & _
                            "          To_Char(s.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi') || '|' || s.ְҵ || '|' || s.��ͥ��ַ || '|' || s.��ͥ�绰 || '|' ||" & vbNewLine & _
                            "          To_Char(s.��������, 'yyyy-mm-dd') || '|' || To_Char(s.ȷ������, 'yyyy-mm-dd') || '|' || s.�������1 || '|' || s.�������2 || '|' ||" & vbNewLine & _
                            "          s.���ע As ��Ϣ" & vbNewLine & _
                            "  From �����걨��¼ S" & vbNewLine & _
                            "  Where s.�ĵ�id = [1]" & IIf(strSender <> "", " And s.������=[4]", "") & "),'||||||||||||||||') ��Ϣ" & vbNewLine & _
                            "From ������Ϣ P, ������ҳ Q, ���ű� A" & vbNewLine & _
                            "Where p.����id = [2] And p.����id = q.����id And [3] Between q.��Ժ���� And" & vbNewLine & _
                            "      Nvl(q.��Ժ����, Sysdate) And q.��Ժ����ID = a.Id" & vbNewLine & _
                            IIf(strPatient <> "", " And P.����=[5]", "") & IIf(lngOutNo = 0, "", " And p.�����=[6]") & IIf(lngInNo = 0, "", " And Q.סԺ��=[7]")
                gstrSQL = gstrSQL & vbNewLine & " Union " & vbNewLine & _
                            "Select '" & rsTemp!ID & "' As ID, p.����id, '" & rsTemp!AntetypeId & "' As �ļ�id, q.id ��ҳID," & vbNewLine & _
                                        "       '" & rsTemp!���� & "' As ����, Decode(" & rsTemp!������Դ & ", 1, '����:', 2, 'סԺ:') || a.���� As ����, Decode(" & rsTemp!������Դ & ", 1, p.�����, 2, p.סԺ��) �����," & vbNewLine & _
                                        "       '" & Format(rsTemp!���ʱ��, "yyyy-mm-dd HH:MM") & "' As �ʱ��, p.����, p.�Ա�, p.����,0 As ����ת��, '" & rsTemp!�༭�� & "' As ���,3 as �༭��ʽ," & vbNewLine & _
                                        "Nvl((Select Decode(Nvl(s.����״̬, 0), -1," & vbNewLine & _
                                        "                 Decode(Sign(To_Date('2014-12-22 10:47:09', 'yyyy-mm-dd hh24:mi:ss') - s.�վ�ʱ��), 1, 0, -1)," & vbNewLine & _
                                        "                 Nvl(s.����״̬, 0)) || '|' || s.�վ��� || '|' || To_Char(s.�վ�ʱ��, 'yyyy-mm-dd hh24:mi') || '|' || s.�վ�˵�� || '|' ||" & vbNewLine & _
                                        "          s.������ || '|' || To_Char(s.����ʱ��, 'yyyy-mm-dd hh24:mi') || '|' || s.���͵�λ || '|' || s.���ͱ�ע || '|' || s.�Ǽ��� || '|' ||" & vbNewLine & _
                                        "          To_Char(s.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi') || '|' || s.ְҵ || '|' || s.��ͥ��ַ || '|' || s.��ͥ�绰 || '|' ||" & vbNewLine & _
                                        "          To_Char(s.��������, 'yyyy-mm-dd') || '|' || To_Char(s.ȷ������, 'yyyy-mm-dd') || '|' || s.�������1 || '|' || s.�������2 || '|' ||" & vbNewLine & _
                                        "          s.���ע As ��Ϣ" & vbNewLine & _
                                        "  From �����걨��¼ S" & vbNewLine & _
                                        "  Where s.�ĵ�id = [1]" & IIf(strSender <> "", " And s.������=[4]", "") & "),'||||||||||||||||') ��Ϣ" & vbNewLine & _
                                        "From ������Ϣ P, ���˹Һż�¼ Q, ���ű� A" & vbNewLine & _
                                        "Where p.����id = [2] And p.����id = q.����id And q.ִ��ʱ��=[3]" & vbNewLine & _
                                        "      And q.ִ�в���ID = a.Id" & vbNewLine & _
                                        IIf(strPatient <> "", " And P.����=[5]", "") & IIf(lngOutNo = 0, "", " And p.�����=[6]")
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(rsTemp!ID), CLng(rsTemp!����ID), CDate(rsTemp!�¼�ʱ��), strSender, strPatient, lngOutNo, lngInNo)
                Do While Not rsData.EOF
                    Set rptRcd = Me.rptList.Records.Add()
                    Set rptItem = rptRcd.AddItem(Val(Split(rsData!��Ϣ, "|")(0)))
                    Select Case rptItem.Value
                    Case 0: rptItem.Icon = 0
                    Case 1: rptItem.Icon = 1
                    Case -1: rptItem.Icon = 2
                    Case 2: rptItem.Icon = 3
                    End Select
                    rptRcd.AddItem CStr(rsData!ID)
                    Select Case Val(Split(rsData!��Ϣ, "|")(0))
                    Case 0: rptRcd.AddItem CStr("a)����д�ļ�������")
                    Case 1: rptRcd.AddItem CStr("b)�ѽ��յļ�������")
                    Case -1: rptRcd.AddItem CStr("c)�Ѿ��յļ�������")
                    Case 2: rptRcd.AddItem CStr("d)�ѱ��͵ļ�������")
                    Case Else: rptRcd.AddItem ""
                    End Select
                    rptRcd.AddItem CStr(rsData!����)
                    rptRcd.AddItem CStr(rsData!����)
                    rptRcd.AddItem CStr("" & rsData!�����)
                    rptRcd.AddItem CStr(rsData!����)
                    rptRcd.AddItem CStr("" & rsData!�Ա�)
                    rptRcd.AddItem CStr("" & rsData!����)
                    rptRcd.AddItem CStr(rsData!�ʱ��)
                    rptRcd.AddItem CStr(rsData!���)
                    rptRcd.AddItem CStr(Mid(rsData!��Ϣ, InStr(rsData!��Ϣ, "|") + 1))
                    rptRcd.AddItem CStr(rsData!����ת��)
                    rptRcd.AddItem CStr(rsData!����ID)
                    rptRcd.AddItem CStr(NVL(rsData!��ҳID, 0))
                    rptRcd.AddItem CStr(rsData!�ļ�ID)
                    rptRcd.AddItem CStr(NVL(rsData!�༭��ʽ, 0))
                    rsData.MoveNext
                Loop
                rsTemp.MoveNext
            Loop
            End If
        End If
    End If
    
    Me.rptList.Populate
    
    If strCurId <> "" Then
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow = False Then
                If CStr(rptRow.Record(mCol.ID).Value) = strCurId Then
                    Set Me.rptList.FocusedRow = rptRow: Exit For
                End If
            End If
        Next
    End If
    If Me.rptList.Rows.Count > 0 Then
        If Me.rptList.FocusedRow Is Nothing Then Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
        If Me.rptList.FocusedRow.GroupRow Then
            strCurId = ""
        Else
            strCurId = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
        End If
    Else
        strCurId = ""
    End If
    zlRefList = Me.rptList.Records.Count
    Exit Function

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    zlRefList = Me.rptList.Records.Count
End Function

Public Sub zlRptPrint(ByVal bytMode As Byte)
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
    If Me.rptList.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '�������ݱ��
    If zlReportToVSFlexGrid(Me.vfgTemp, Me.rptList) = False Then Exit Sub
    
    '-------------------------------------------------
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vfgTemp
    objPrint.Title.Text = "�����ļ��嵥"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

'-------------------------------------------------------
'���ܣ�  ����Ԥ������ӡ
'������  blnPreview  :�Ƿ���Ԥ��ģʽ
'-------------------------------------------------------
Private Sub zlEPRPrint(blnPreview As Boolean)
Dim frmPrint As frmPrintPreview, ObjTabEprView As Object
Dim rsTemp As New ADODB.Recordset
    If mstrCurId = "" Then Exit Sub
    Err = 0: On Error GoTo errHand
    If IsNumeric(mstrCurId) Then
        gstrSQL = "Select l.������Դ, l.����id, l.��ҳid,l.�༭��ʽ, f.ҳ�� From ���Ӳ�����¼ l, �����ļ��б� f Where l.�ļ�id = f.Id And l.Id = [1]"
        If mblnCurMoved Then
            gstrSQL = Replace(gstrSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(mstrCurId))
        With rsTemp
            If .RecordCount <= 0 Then MsgBox "�ü�����������Ѿ����ٴ�ɾ����", vbExclamation, gstrSysName: Exit Sub
            If rsTemp!�༭��ʽ = 0 Then
                Set frmPrint = New frmPrintPreview
                Select Case !������Դ
                Case 1
                    frmPrint.DoMultiDocPreview Me, cpr���ﲡ��, !����ID, !��ҳID, 5, !ҳ��, CLng(mstrCurId), Not blnPreview, , , mblnCurMoved
                Case 2
                    frmPrint.DoMultiDocPreview Me, cprסԺ����, !����ID, !��ҳID, 5, !ҳ��, CLng(mstrCurId), Not blnPreview, , , mblnCurMoved
                End Select
                Unload frmPrint
                Set frmPrint = Nothing
            ElseIf rsTemp!�༭��ʽ = 1 Then
                Set ObjTabEprView = DynamicCreate("zlTableEPR.cTableEPR", "��ӡ�����", True)
                Call ObjTabEprView.InitTableEPR(gcnOracle, glngSys, gstrDbOwner)
                Call ObjTabEprView.InitOpenEPR(Me, cprEM_�޸�, cprET_���������, CLng(mstrCurId), False, 0, !������Դ)
                ObjTabEprView.zlPrintDoc Me, blnPreview, ""
            ElseIf rsTemp!�༭��ʽ = 2 Then
                mobjInfection.PrintDoc Me, !����ID, !��ҳID, CLng(mstrCurId), ""
            End If
        End With
    Else
        If Not mobjRichEMR Is Nothing Then Call mobjRichEMR.zlPrintDoc(blnPreview)
    End If
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'-------------------------------------------------------
'����Ϊ�ؼ��¼�����
'-------------------------------------------------------

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim cbrControl As CommandBarControl, strInfo As String, bytEdit As Byte
    If mblnCurMoved And (Control.ID = conMenu_File_Open Or Control.ID = conMenu_Edit_Reuse Or Control.ID = conMenu_Edit_Send Or Control.ID = conMenu_Edit_Untread) Then
        MsgBox "�ò��˵ı��������Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                        "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Sub
    End If

    If rptList.FocusedRow Is Nothing Then
        bytEdit = 0
    ElseIf rptList.FocusedRow.GroupRow = True Then
        bytEdit = 0
    Else
        bytEdit = Val(rptList.FocusedRow.Record.Item(mCol.�༭��ʽ).Value)
    End If
            
    Select Case Control.ID
    Case conMenu_File_Open
        If bytEdit = 0 Then
            Dim f As New frmEPRView
            f.ShowMe Me, CLng(mstrCurId)
        ElseIf bytEdit = 3 Then
            '�²����༭��
            Call mobjRichEMR.zlViewDoc(Me, "����", "")
        End If
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlEPRPrint(True)
    Case conMenu_File_Print: Call zlEPRPrint(False)
    Case conMenu_File_RowPrint: Call zlRptPrint(1)
    Case conMenu_File_Parameter
        Call frmDiseaseReportSet.ShowMe(Me, InStr(1, mstrPrivs, "��Χ����") > 0, mstrFiles, mintDates, mstrDateFrom, mstrDateTo)
        Call zlRefList
    Case conMenu_File_Exit: Unload Me
    Case conMenu_Edit_Audit '��˲���
        '���������ģʽ
        If bytEdit = 0 Then
            Dim frmAudit As Form, bFindedAudit As Boolean
            For Each frmAudit In Forms
                If frmAudit.Name = "frmMain" Then
                    If frmAudit.Document.EPRPatiRecInfo.ID = CLng(mstrCurId) Then
                        frmAudit.Show
                        bFindedAudit = True
                    End If
                End If
            Next
            If bFindedAudit = False Then
                Set mobjDoc = New cEPRDocument
                mobjDoc.InitEPRDoc cprEM_�޸�, cprET_���������, CLng(mstrCurId), cprPF_סԺ
                mobjDoc.ShowEPREditor Me
            End If
        ElseIf bytEdit = 3 Then
            '�²����༭��
            Dim objAudit As Object
            Set objAudit = DynamicCreate("zlRichEMR.clsDockEMR", "�°没��", False)
            Call objAudit.Init(gobjEmr, gcnOracle, glngSys)
            Call objAudit.zlRefresh(rptList.FocusedRow.Record.Item(mCol.����ID).Value, rptList.FocusedRow.Record.Item(mCol.��ҳID).Value, glngDeptId, 0, IIf(InStr(rptList.FocusedRow.Record.Item(mCol.����).Value, "����") > 0, 1, 2))
            Call objAudit.EditDoc(mstrCurId)
        End If
    Case conMenu_Edit_Reuse
        'strInfo=����|����|����|�Ա�|����|�����|���|�ʱ��|����ID|��ҳID
        With rptList
            strInfo = .FocusedRow.Record.Item(mCol.����).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.����).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.����).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.�Ա�).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.����).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.�����).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.���).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.�ʱ��).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.����ID).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.��ҳID).Value
        End With
        If Not IsNumeric(mstrCurId) Then
            '�²��� strInfo=strInfo & |���ഫȾ��|���ഫȾ��|���ഫȾ��|��������|��������2
            strInfo = strInfo & "|" & rptList.Tag
        End If
        If frmDiseaseReportIncept.ShowMe(Me, mstrCurId, strInfo) Then Call zlRefList(mstrCurId)
    Case conMenu_Edit_Send
        'strInfo=����|����|����|�Ա�|����|�����|���|�ʱ��|����ID|��ҳID
        With rptList
            strInfo = .FocusedRow.Record.Item(mCol.����).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.����).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.����).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.�Ա�).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.����).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.�����).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.���).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.�ʱ��).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.����ID).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.��ҳID).Value
            strInfo = strInfo & "|" & .FocusedRow.Record.Item(mCol.�ļ�ID).Value
        End With
        If frmDiseaseReportSend.ShowMe(Me, mstrCurId, strInfo, Me.vfgInfo.TextMatrix(5, 1), Me.vfgInfo.TextMatrix(6, 1)) Then Call zlRefList(mstrCurId)
    Case conMenu_Edit_Untread
        Dim strMsg As String
        Select Case mintState
        Case 1:  strMsg = "���ȡ���ü�������ġ����մ�����"
        Case -1: strMsg = "���ȡ���ü�������ġ��ܾ�������"
        Case 2:  strMsg = "���ȡ���ü�������ġ��걨�Ǽǡ���"
        Case Else: Exit Sub
        End Select
        If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        gstrSQL = "Zl_�����걨��¼_Untread('" & mstrCurId & "')"
        Err = 0: On Error GoTo errHand
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Call zlRefList(mstrCurId)
    Case conMenu_Edit_Compend '��Ӧ��Ϣ����
        frmDiseaseReportRela.Show 1, Me
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.STYLE = IIf(cbrControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_Refresh
        Call zlRefList(mstrCurId)
    Case conMenu_View_Find
        Dim strSender As String, strPatient As String, lngOutNo As Long, lngInNo As Long
        If frmDiseaseReportFind.ShowMe(Me, mintDates, mstrDateFrom, mstrDateTo, strSender, strPatient, lngOutNo, lngInNo) Then
            Call zlRefList(mstrCurId, strSender, strPatient, lngOutNo, lngInNo)
        End If
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Forum '������̳
        Call zlWebForum(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    Case Else
        'ִ�з�������ǰģ��ı���
        Dim lng����ID As Long
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            If rptList.SelectedRows.Count > 0 Then
                If Not rptList.SelectedRows(0).GroupRow Then
                    lng����ID = Val(rptList.SelectedRows(0).Record(mCol.ID).Value)
                End If
            End If
            If lng����ID <> 0 Then
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "����ID=" & lng����ID)
            Else
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
            End If
        End If
    End Select
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Open
         Control.Enabled = (mstrCurId <> "")
         If Control.Enabled Then Control.Enabled = InStr("0,3", Me.rptList.FocusedRow.Record.Item(mCol.�༭��ʽ).Value) > 0 'Ŀǰ��֧���ϲ��������ܲ���
    Case conMenu_File_Preview
         Control.Enabled = (mstrCurId <> "")
         If Control.Enabled Then Control.Enabled = InStr("0,1,3", Me.rptList.FocusedRow.Record.Item(mCol.�༭��ʽ).Value) > 0 'Ŀǰ��֧���ϲ��������ܲ����������
    Case conMenu_File_Print
         Control.Enabled = (mstrCurId <> "")
         If Control.Enabled Then Control.Enabled = InStr("0,1,3", Me.rptList.FocusedRow.Record.Item(mCol.�༭��ʽ).Value) > 0 'Ŀǰ��֧���ϲ��������ܲ����������
    Case conMenu_File_RowPrint: Control.Enabled = (Me.rptList.Records.Count <> "")
    Case conMenu_Edit_Audit
        Control.Visible = (mstrCurId <> "")
        If Control.Visible Then Control.Visible = InStr("0,3", Me.rptList.FocusedRow.Record.Item(mCol.�༭��ʽ).Value) > 0 'Ŀǰ��֧���ϲ��������ܲ���
        Control.Enabled = (InStr(1, mstrPrivs, "��������") > 0)
        If Control.Enabled Then Control.Enabled = (mstrCurId <> "")
        If Control.Enabled Then Control.Enabled = (mintState = 0)
    Case conMenu_Edit_Reuse
        Control.Enabled = (InStr(1, mstrPrivs, "����") > 0)
        If Control.Enabled Then Control.Enabled = (mstrCurId <> "" And mintState = 0)
    Case conMenu_Edit_Send
        Control.Enabled = (InStr(1, mstrPrivs, "����") > 0)
        If Control.Enabled Then Control.Enabled = (mstrCurId <> "" And mintState = 1)
    Case conMenu_Edit_Untread
        Control.Enabled = (InStr(1, mstrPrivs, "����") > 0)
        If Control.Enabled Then Control.Enabled = (mstrCurId <> "" And mintState <> 0)
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).STYLE = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    Case conMenu_View_Refresh:  Control.Enabled = (Trim(mstrFiles) <> "")
    Case conMenu_View_Find::  Control.Enabled = (Trim(mstrFiles) <> "")
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_Reports
        Item.Handle = Me.rptList.hWnd
    Case conPane_Preview
        If mfrmPreview Is Nothing Then Set mfrmPreview = New frmDockEPRContent
        Item.Handle = mfrmPreview.hWnd
    Case conPane_AppInfo
        Item.Handle = Me.vfgInfo.hWnd
    End Select
End Sub

Private Sub Form_Load()
Dim cbrMenuBar As CommandBarPopup
Dim cbrControl As CommandBarControl
Dim cbrToolBar As CommandBar
Dim rptCol As ReportColumn
Dim lngCount As Long
    '-----------------------------------------------------
    'Ȩ�����ƴ����ƣ�����ͬʱ��������ģ�������gstrPrivs�仯�����¿�����Ч
    mstrPrivs = gstrPrivs
    mstrFiles = Trim(GetSetting("ZLSOFT", App.EXEName, "�����걨�ļ���Χ", ""))
    mintDates = Val(GetSetting("ZLSOFT", App.EXEName, "�����걨�������", 0))
    If mintDates = 0 Then mintDates = 7: Call SaveSetting("ZLSOFT", App.EXEName, "�����걨�������", mintDates)
    
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    '�˵�����
    Me.cbsThis.ActiveMenuBar.Title = "�˵�"
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "��(&O)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_RowPrint, "�嵥��ӡ(&L)��"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&M)��"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "�޶�(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "����(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Send, "����(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "�ջ�(&B)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Compend, "��Ӧ��Ϣ����(&B)")
        cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Find, "����(&F)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False  '����
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With
    
    '�����
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("O"), conMenu_File_Open
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("F"), conMenu_View_Find
        .Add 0, vbKeyF12, conMenu_File_Parameter
        .Add FCONTROL, Asc("A"), conMenu_Edit_Reuse
        .Add FCONTROL, Asc("T"), conMenu_Edit_Send
        .Add FCONTROL, Asc("U"), conMenu_Edit_Untread
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F6, conMenu_View_Jump
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '���ò����ò˵�
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_RowPrint
        .AddHiddenCommand conMenu_View_Refresh
    End With
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "�޶�")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "����")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Send, "����")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "����")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.STYLE = xtpButtonIconAndCaption
    Next
    
    '��ȡ��������ģ��ı���:��Ϊ��һ���Զ�ȡ,ȫ�ֱ�������
    '---------------------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    
    '-----------------------------------------------------
    '���ôʾ���ʾͣ������
    If mfrmPreview Is Nothing Then Set mfrmPreview = New frmDockEPRContent
    If Not gobjEmr Is Nothing Then
        Set mobjRichEMR = DynamicCreate("zlRichEMR.clsDockContent", "�°没��", False)
        If Not mobjRichEMR Is Nothing Then Call mobjRichEMR.Init(gobjEmr, gcnOracle, glngSys, 0)
    End If

    Set mobjInfection = DynamicCreate("zlDisReportCard.clsDisReportCard", "��Ⱦ�����濨", True)
    If Not mobjInfection Is Nothing Then
        mobjInfection.Init gcnOracle, glngSys
    End If
    
    Dim panThis As Pane, panChild As Pane
    Set panThis = dkpMan.CreatePane(conPane_Reports, 400, 200, DockLeftOf, Nothing)
    panThis.Title = "�����б�": panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Set panChild = dkpMan.CreatePane(conPane_Preview, 400, 200, DockBottomOf, Nothing)
    panChild.Title = "��������": panChild.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Set panThis = dkpMan.CreatePane(conPane_AppInfo, 200, 400, DockRightOf, Nothing)
    panThis.Title = "������Ϣ": panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    
    '-----------------------------------------------------
    With Me.rptList
        Set rptCol = .Columns.Add(mCol.ͼ��, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.״̬, "״̬", 90, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.����, "����", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.����, "����", 110, True): rptCol.Editable = False: rptCol.Groupable = True
        Set rptCol = .Columns.Add(mCol.�����, "����&סԺ��", 75, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.����, "����", 50, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.�Ա�, "�Ա�", 40, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.����, "����", 40, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.�ʱ��, "�ʱ��", 100, True): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mCol.���, "���", 50, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.��Ϣ, "��Ϣ", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.����ת��, "����ת��", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.����ID, "����ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.��ҳID, "��ҳID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.�ļ�ID, "�ļ�ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        .GroupsOrder.Add .Columns.Find(mCol.״̬)
        .GroupsOrder(0).SortAscending = True
        
        .SetImageList Me.imgList
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With
    
    '-----------------------------------------------------
    '����ָ�
    Call RestoreWinState(Me, App.ProductName)
    '-----------------------------------------------------
    '����װ��
    If mstrFiles = "" Then
        Me.stbThis.Panels(2).Text = "δ���ñ�����վ�ļ������淶Χ"
    Else
        lngCount = zlRefList()
        Me.stbThis.Panels(2).Text = "����" & lngCount & "�ݼ�������"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Unload mfrmPreview
    Set mfrmPreview = Nothing
    Set mobjDoc = Nothing
    Unload mobjRichEMR.zlGetForm
    Set mobjRichEMR.zlGetForm = Nothing
    Set mobjRichEMR = Nothing
    Set mobjInfection = Nothing
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
Dim cbrPopupBar As CommandBar
Dim cbrPopupItem As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrControl As CommandBarControl
    
    If Button <> vbRightButton Then Exit Sub
    If Me.cbsThis.ActiveMenuBar.Controls(2).Visible = False Then Exit Sub

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = Me.cbsThis.Add("�����˵�", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
    Next
    cbrPopupBar.ShowPopup
End Sub

Private Sub rptList_SelectionChanged()
    Dim strInfo As String, aryInfo() As String
    
    mstrContent = "": rptList.Tag = ""
    With Me.rptList
        If .FocusedRow Is Nothing Then
            mstrCurId = "": mintState = 0: strInfo = "": mblnCurMoved = False
        ElseIf .FocusedRow.GroupRow = True Then
            mstrCurId = "": mintState = 0: strInfo = "": mblnCurMoved = False
        Else
            mstrCurId = .FocusedRow.Record.Item(mCol.ID).Value
            mintState = .FocusedRow.Record.Item(mCol.ͼ��).Value
            strInfo = .FocusedRow.Record.Item(mCol.��Ϣ).Value
            mblnCurMoved = (.FocusedRow.Record.Item(mCol.����ת��).Value = 1)
        End If
    End With
    
    aryInfo = Split(strInfo, "|")
    With Me.vfgInfo
        .Clear
        .ColWidth(0) = 900
        Select Case mintState
        Case 0
            If strInfo <> "" Then .Rows = 1: .TextMatrix(0, 0) = "������": .TextMatrix(0, 1) = "�ȴ����ա�"
        Case 1
            .Rows = 11
            .TextMatrix(0, 0) = "ְҵ": .TextMatrix(0, 1) = aryInfo(9)
            .TextMatrix(1, 0) = "��ͥ��ַ": .TextMatrix(1, 1) = aryInfo(10)
            .TextMatrix(2, 0) = "��ͥ�绰": .TextMatrix(2, 1) = aryInfo(11)
            .TextMatrix(3, 0) = "��������": .TextMatrix(3, 1) = aryInfo(12)
            .TextMatrix(4, 0) = "ȷ������": .TextMatrix(4, 1) = aryInfo(13)
            .TextMatrix(5, 0) = "�������1": .TextMatrix(5, 1) = aryInfo(14)
            .TextMatrix(6, 0) = "�������2": .TextMatrix(6, 1) = aryInfo(15)
            .TextMatrix(7, 0) = "���ע": .TextMatrix(7, 1) = aryInfo(16)
            
            .TextMatrix(8, 0) = "������": .TextMatrix(8, 1) = aryInfo(0)
            .TextMatrix(9, 0) = "����ʱ��": .TextMatrix(9, 1) = aryInfo(1)
            .TextMatrix(10, 0) = "����˵��": .TextMatrix(10, 1) = aryInfo(2)
        Case -1
            .Rows = 3
            .TextMatrix(0, 0) = "������": .TextMatrix(0, 1) = aryInfo(0)
            .TextMatrix(1, 0) = "����ʱ��": .TextMatrix(1, 1) = aryInfo(1)
            .TextMatrix(2, 0) = "����ԭ��": .TextMatrix(2, 1) = aryInfo(2)
        Case 2
            .Rows = 17
            .TextMatrix(0, 0) = "ְҵ": .TextMatrix(0, 1) = aryInfo(9)
            .TextMatrix(1, 0) = "��ͥ��ַ": .TextMatrix(1, 1) = aryInfo(10)
            .TextMatrix(2, 0) = "��ͥ�绰": .TextMatrix(2, 1) = aryInfo(11)
            .TextMatrix(3, 0) = "��������": .TextMatrix(3, 1) = aryInfo(12)
            .TextMatrix(4, 0) = "ȷ������": .TextMatrix(4, 1) = aryInfo(13)
            .TextMatrix(5, 0) = "�������1": .TextMatrix(5, 1) = aryInfo(14)
            .TextMatrix(6, 0) = "�������2": .TextMatrix(6, 1) = aryInfo(15)
            .TextMatrix(7, 0) = "���ע": .TextMatrix(7, 1) = aryInfo(16)
            
            .TextMatrix(8, 0) = "������": .TextMatrix(8, 1) = aryInfo(0)
            .TextMatrix(9, 0) = "����ʱ��": .TextMatrix(9, 1) = aryInfo(1)
            .TextMatrix(10, 0) = "����˵��": .TextMatrix(10, 1) = aryInfo(2)
            .TextMatrix(11, 0) = "������": .TextMatrix(11, 1) = aryInfo(3)
            .TextMatrix(12, 0) = "����ʱ��": .TextMatrix(12, 1) = aryInfo(4)
            .TextMatrix(13, 0) = "���͵�λ": .TextMatrix(13, 1) = aryInfo(5)
            .TextMatrix(14, 0) = "���ͱ�ע": .TextMatrix(14, 1) = aryInfo(6)
            .TextMatrix(15, 0) = "�Ǽ���": .TextMatrix(15, 1) = aryInfo(7)
            .TextMatrix(16, 0) = "�Ǽ�ʱ��": .TextMatrix(16, 1) = aryInfo(8)
        End Select
    End With
    
    On Error Resume Next
    If IsNumeric(mstrCurId) Then
        dkpMan.FindPane(conPane_Preview).Handle = mfrmPreview.hWnd
        Call mfrmPreview.zlRefresh(CLng(mstrCurId), "", , mblnCurMoved, , NVL(rptList.FocusedRow.Record.Item(mCol.�༭��ʽ).Value, 0))
    ElseIf mstrCurId <> "" Then
        dkpMan.FindPane(conPane_Preview).Handle = mobjRichEMR.zlGetForm.hWnd
        Call mobjRichEMR.zlShowDoc(mstrCurId, "")
        Call mobjRichEMR.zlGetForm.DocContent.SaveToXML(mstrContent, False)
        
        If mstrContent <> "" Then
            Dim xmldom As Object, xmlNode As Object, xmlEle As Object, sxpath As String, lngItem As Long, strCDATA As String, strItem As String
            Set xmldom = CreateObject("Msxml2.DOMDocument.6.0")
            Set xmlEle = CreateObject("Msxml2.DOMDocument.6.0")
            Call xmldom.loadXML(mstrContent)
            
            lngItem = 0: strItem = ""
            sxpath = "/zlxml/document/e_enum[contains(@title,""���ഫȾ��"")]"
            Call xmlEle.loadXML(xmldom.selectSingleNode(sxpath).xml)
            Set xmlNode = xmlEle.selectSingleNode("/e_enum/enumvalues/element")
            If Not xmlNode Is Nothing Then
                lngItem = Val(xmlNode.Text)
                If xmldom.selectSingleNode(sxpath).firstChild.nodeType = NODE_CDATA_SECTION Then
                    strCDATA = xmldom.selectSingleNode(sxpath).firstChild.nodeValue
                    strCDATA = Replace(strCDATA, "rangexml='", "")
                    strCDATA = Mid(strCDATA, 1, Len(strCDATA) - 1)
                    Call xmlEle.loadXML(strCDATA)
                    strItem = xmlEle.selectSingleNode("/root/item[" & lngItem & "]/meaning").Text
                End If
            End If
            strInfo = strItem
            
            lngItem = 0: strItem = ""
            sxpath = "/zlxml/document/e_enum[contains(@title,""���ഫȾ��"")]"
            Call xmlEle.loadXML(xmldom.selectSingleNode(sxpath).xml)
            Set xmlNode = xmlEle.selectSingleNode("/e_enum/enumvalues/element")
            If Not xmlNode Is Nothing Then
                lngItem = Val(xmlNode.Text)
                If xmldom.selectSingleNode(sxpath).firstChild.nodeType = NODE_CDATA_SECTION Then
                    strCDATA = xmldom.selectSingleNode(sxpath).firstChild.nodeValue
                    strCDATA = Replace(strCDATA, "rangexml='", "")
                    strCDATA = Mid(strCDATA, 1, Len(strCDATA) - 1)
                    Call xmlEle.loadXML(strCDATA)
                    strItem = xmlEle.selectSingleNode("/root/item[" & lngItem & "]/meaning").Text
                End If
            End If
            strInfo = strInfo & "|" & strItem
            
            lngItem = 0: strItem = ""
            sxpath = "/zlxml/document/e_enum[contains(@title,""���ഫȾ��"")]"
            Call xmlEle.loadXML(xmldom.selectSingleNode(sxpath).xml)
            Set xmlNode = xmlEle.selectSingleNode("/e_enum/enumvalues/element")
            If Not xmlNode Is Nothing Then
                lngItem = Val(xmlNode.Text)
                If xmldom.selectSingleNode(sxpath).firstChild.nodeType = NODE_CDATA_SECTION Then
                    strCDATA = xmldom.selectSingleNode(sxpath).firstChild.nodeValue
                    strCDATA = Replace(strCDATA, "rangexml='", "")
                    strCDATA = Mid(strCDATA, 1, Len(strCDATA) - 1)
                    Call xmlEle.loadXML(strCDATA)
                    strItem = xmlEle.selectSingleNode("/root/item[" & lngItem & "]/meaning").Text
                End If
            End If
            strInfo = strInfo & "|" & strItem
            
            lngItem = 0: strItem = ""
            sxpath = "/zlxml/document/e_enum[contains(@title,""��������"")][1]"
            Call xmlEle.loadXML(xmldom.selectSingleNode(sxpath).xml)
            Set xmlNode = xmlEle.selectSingleNode("/e_enum/enumvalues/element")
            If Not xmlNode Is Nothing Then
                lngItem = Val(xmlNode.Text)
                If xmldom.selectSingleNode(sxpath).firstChild.nodeType = NODE_CDATA_SECTION Then
                    strCDATA = xmldom.selectSingleNode(sxpath).firstChild.nodeValue
                    strCDATA = Replace(strCDATA, "rangexml='", "")
                    strCDATA = Mid(strCDATA, 1, Len(strCDATA) - 1)
                    Call xmlEle.loadXML(strCDATA)
                    strItem = xmlEle.selectSingleNode("/root/item[" & lngItem & "]/meaning").Text
                End If
            End If
            strInfo = strInfo & "|" & strItem
            
            lngItem = 0: strItem = ""
            sxpath = "/zlxml/document/e_enum[contains(@title,""��������"")][2]"
            Call xmlEle.loadXML(xmldom.selectSingleNode(sxpath).xml)
            Set xmlNode = xmlEle.selectSingleNode("/e_enum/enumvalues/element")
            If Not xmlNode Is Nothing Then
                lngItem = Val(xmlNode.Text)
                If xmldom.selectSingleNode(sxpath).firstChild.nodeType = NODE_CDATA_SECTION Then
                    strCDATA = xmldom.selectSingleNode(sxpath).firstChild.nodeValue
                    strCDATA = Replace(strCDATA, "rangexml='", "")
                    strCDATA = Mid(strCDATA, 1, Len(strCDATA) - 1)
                    Call xmlEle.loadXML(strCDATA)
                    strItem = xmlEle.selectSingleNode("/root/item[" & lngItem & "]/meaning").Text
                End If
            End If
            strInfo = strInfo & "|" & strItem
            
        End If
        rptList.Tag = strInfo
    End If
End Sub




