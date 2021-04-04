VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmReportHistory 
   Caption         =   "�����޶���ʷ"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9735
   Icon            =   "frmReportHistory.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   9735
   StartUpPosition =   1  '����������
   Begin RichTextLib.RichTextBox rtxtReport 
      Height          =   4455
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   7858
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmReportHistory.frx":0CCA
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   120
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmReportHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngAdviceID As Long        'ҽ��ID
Private mlngPatientId As Long       '����ID
Private mlngCur����ID As Long       '����ID
Private mlngReportID As Long        '����ID
Private mlngMode As Long            '����鿴״̬��0-�޶�״̬��1-����״̬
Private mintReportCount As Integer  '��ʷ���������
Private mlngViewReportID As Long    '��ǰ�鿴�ı���ID
Private mlngViewAdviceID As Long    '��ǰ�鿴��ҽ��ID
Private mstrOffset As String        '��ǰ����ߵ�����

Private mobjReport As zlRichEPR.cDockReport    '�������



Public Sub zlShowMe(frmParent As Object, lngAdviceID As Long, lngReportID As Long)
    mlngAdviceID = lngAdviceID
    mlngReportID = lngReportID
    mlngViewReportID = mlngReportID
    mlngViewAdviceID = mlngAdviceID
    Me.Show 0, frmParent
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    Select Case control.ID
        Case conMenu_PacsReport_Mode_Orig                   'ԭʼ״̬
            If mlngMode <> 0 Then
                ShowModeOrig mlngViewReportID, mlngViewAdviceID
            End If
            mlngMode = 0
            Me.cbrMain.FindControl(, conMenu_PacsReport_Mode_Clear, , True).Checked = False
            control.Checked = True
        Case conMenu_PacsReport_Mode_Clear                  '����״̬
            If mlngMode <> 1 Then
                ShowModeClear mlngViewReportID, mlngViewAdviceID
            End If
            mlngMode = 1
            Me.cbrMain.FindControl(, conMenu_PacsReport_Mode_Orig, , True).Checked = False
            control.Checked = True
        Case conMenu_File_Preview                           '����Ԥ��
            If mlngViewReportID = 0 Then Exit Sub
            mobjReport.zlRefresh 0, 0
            mobjReport.zlRefresh mlngViewAdviceID, UserInfo.����ID
            mobjReport.zlExecuteCommandBars control
        Case conMenu_File_Exit                              '   �˳�
                Unload Me
        Case Else
            ShowHistory control.ID
    End Select
End Sub

Private Sub cbrMain_Resize()
    Dim iLeft As Long, iTop As Long, iRight As Long, iBottom As Long
    cbrMain.GetClientRect iLeft, iTop, iRight, iBottom
    rtxtReport.Left = iLeft
    rtxtReport.Top = iTop
    rtxtReport.Width = Abs(iRight - iLeft)
    rtxtReport.Height = Abs(iBottom - iTop)
End Sub

Private Sub ShowHistory(iIndex As Integer)
    Dim lngReportID As Long
    Dim lngAdviceID As Long
    Dim strTemp As String
    
    If iIndex > conMenu_PacsReport_History_Times And iIndex <= conMenu_PacsReport_History_Times + mintReportCount Then
        strTemp = Me.cbrMain.FindControl(, iIndex, , True).DescriptionText
        If InStr(strTemp, "-") <> 0 Then
            lngReportID = Val(Split(strTemp, "-")(1))
            lngAdviceID = Val(Split(strTemp, "-")(0))
            mlngViewReportID = lngReportID
            mlngViewAdviceID = lngAdviceID
            If mlngMode = 0 Then
                Call ShowModeOrig(mlngViewReportID, mlngViewAdviceID)
            ElseIf mlngMode = 1 Then
                Call ShowModeClear(mlngViewReportID, mlngViewAdviceID)
            End If
        End If
    End If
End Sub

Private Sub ShowTitle(lngReportID As Long, lngAdviceID As Long)
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strSeparator1 As String
    Dim strSeparator2 As String
    Dim lngStart As Long
    Dim strTitle As String
    Dim strWriter As String
    Dim lng����ID As Long
    Dim lng��ҳID As Long
    Dim intӤ�� As Integer
    Dim rsBaby As ADODB.Recordset
    
    If lngReportID = 0 Then Exit Sub
    
    strSeparator1 = mstrOffset & "==================================================" & vbCrLf
    strSeparator2 = mstrOffset & "-------------------" & vbCrLf
    
    strSQL = "Select a.����,a.����,b.����ʱ��,b.ҽ������,a.������,a.������,nvl(b.Ӥ��,0) as Ӥ��,a.�������� as ���ʱ��, " _
            & "b.����ID, nvl(b.��ҳID,0) as ��ҳID From Ӱ�����¼ a,����ҽ����¼ b Where a.ҽ��id = b.Id And b.Id = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngAdviceID)
    If rsTemp.EOF = True Then Exit Sub
    
    strTitle = mstrOffset & nvl(rsTemp!ҽ������) & vbCrLf
    
    lngStart = Len(rtxtReport.Text)
    rtxtReport.SelStart = lngStart
    rtxtReport.SelLength = 0
    rtxtReport.SelText = strTitle
    
    rtxtReport.SelStart = lngStart
    rtxtReport.SelLength = Len(strTitle)
    rtxtReport.SelFontName = "����"
    rtxtReport.SelFontSize = 16
    rtxtReport.SelBold = True
    rtxtReport.SelColor = vbBlue
    
    'Ӥ����������Ҫ������ʾ
    If rsTemp!Ӥ�� = 0 Then
        strWriter = vbCrLf & mstrOffset & "������" & nvl(rsTemp!����) & "      ���ţ�" & nvl(rsTemp!����) & vbCrLf _
               & mstrOffset & "�����ˣ�" & nvl(rsTemp!������) & "      ����ˣ�" & nvl(rsTemp!������) & vbCrLf _
               & mstrOffset & "����ʱ�䣺 " & nvl(rsTemp!����ʱ��) & "      ���ʱ�䣺" & nvl(rsTemp!���ʱ��) & vbCrLf
    Else
        lng����ID = rsTemp!����ID
        lng��ҳID = rsTemp!��ҳID
        intӤ�� = rsTemp!Ӥ��
        strSQL = "Select Decode(a.Ӥ������,Null,b.����||'֮��'||Trim(To_Char(a.���,'9')),a.Ӥ������) As Ӥ������,Ӥ���Ա�,����ʱ�� From ������������¼ a,������Ϣ b Where a.����id=[1] And a.��ҳid=[2] And a.����id=b.����id And a.���=[3]"
        Set rsBaby = zlDatabase.OpenSQLRecord(strSQL, "����Ӥ����Ϣ", lng����ID, lng��ҳID, intӤ��)
        
        strWriter = vbCrLf & mstrOffset & "������" & rsBaby!Ӥ������ & "      ���ţ�" & nvl(rsTemp!����) & vbCrLf _
               & mstrOffset & "�����ˣ�" & nvl(rsTemp!������) & "      ����ˣ�" & nvl(rsTemp!������) & vbCrLf _
               & mstrOffset & "����ʱ�䣺 " & nvl(rsTemp!����ʱ��) & "      ���ʱ�䣺" & nvl(rsTemp!���ʱ��) & vbCrLf
    
    End If
'    '������Ϣ
'    strSQL = "Select �������� From ���Ӳ�����¼  Where Id =  [1] "
'    Set rsTemp = OpenSQLRecord(strSQL, Me.Caption, lngReportID)
'    If rsTemp.EOF = True Then Exit Sub
'
'    strTitle = mstrOffset & Nvl(rsTemp!��������) & vbCrLf
'
'    lngStart = Len(rtxtReport.Text)
'    rtxtReport.SelStart = lngStart
'    rtxtReport.SelLength = 0
'    rtxtReport.SelText = strTitle
'
'    rtxtReport.SelStart = lngStart
'    rtxtReport.SelLength = Len(strTitle)
'    rtxtReport.SelFontName = "����"
'    rtxtReport.SelFontSize = 14
'    rtxtReport.SelBold = False
'    rtxtReport.SelColor = vbBlue
    
    '��ʾ������
    strWriter = strWriter
    
    lngStart = Len(rtxtReport.Text)
    rtxtReport.SelStart = lngStart
    rtxtReport.SelLength = 0
    rtxtReport.SelText = strWriter
    
    rtxtReport.SelStart = lngStart
    rtxtReport.SelLength = Len(strWriter)
    rtxtReport.SelFontName = "����"
    rtxtReport.SelFontSize = 14
    rtxtReport.SelBold = False
    rtxtReport.SelColor = vbBlue
    
    '��ʾ����
    lngStart = Len(rtxtReport.Text)
    rtxtReport.SelStart = lngStart
    rtxtReport.SelLength = 0
    rtxtReport.SelText = strSeparator1
    
    rtxtReport.SelStart = lngStart
    rtxtReport.SelLength = Len(strSeparator1)
    rtxtReport.SelFontName = "����"
    rtxtReport.SelFontSize = 14
    rtxtReport.SelBold = False
    rtxtReport.SelColor = vbBlack
    
'    'ǩ����Ϣ
'    strSQL = "Select �����ı� As ǩ���� ,Ҫ������ As ǩ��ǰ׺,������ From ���Ӳ������� b Where  b.��������=8 And �ļ�ID= [1] Order By ������ "
'    Set rsTemp = OpenSQLRecord(strSQL, Me.Caption, lngReportID)
'    If rsTemp.EOF = True Then Exit Sub
'
'    strTitle = mstrOffset & "ǩ���ˣ�" & Nvl(rsTemp!ǩ��ǰ׺) & Nvl(rsTemp!ǩ����) & vbCrLf
'    rsTemp.MoveNext
'    While Not rsTemp.EOF
'        strTitle = strTitle & mstrOffset & "        " & Nvl(rsTemp!ǩ��ǰ׺) & Nvl(rsTemp!ǩ����) & vbCrLf
'        rsTemp.MoveNext
'    Wend
'    strTitle = strTitle & strSeparator1
'
'    lngStart = Len(rtxtReport.Text)
'    rtxtReport.SelStart = lngStart
'    rtxtReport.SelLength = 0
'    rtxtReport.SelText = strTitle
'
'    rtxtReport.SelStart = lngStart
'    rtxtReport.SelLength = Len(strTitle)
'    rtxtReport.SelFontName = "����"
'    rtxtReport.SelFontSize = 14
'    rtxtReport.SelBold = False
'    rtxtReport.SelColor = vbBlue
End Sub

Private Sub Form_Load()
    
    mlngMode = 1
    mintReportCount = 0
    mstrOffset = "  "
    Set mobjReport = New zlRichEPR.cDockReport      '���Ӳ�������
    
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitCommandBars '��ʼ���˵�
    
    If mlngReportID = 0 Then    '��ǰ����û�б��棬ֱ����ʾ�����һ����ʷ����
        If mintReportCount >= 1 Then
            ShowHistory conMenu_PacsReport_History_Times + mintReportCount
        End If
    Else
        ShowModeClear mlngViewReportID, mlngViewAdviceID
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Unload mobjReport.zlGetForm        '���Ӳ�������
    '���洰��λ��
    Call SaveWinState(Me, App.ProductName)
End Sub


Private Sub InitCommandBars()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim cbrPopControl As CommandBarControl
    Dim strSQL  As String
    Dim strSQLBak As String
    Dim rsTemp As ADODB.Recordset
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    Set Me.cbrMain.Icons = zlCommFun.GetPubIcons
    
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.Visible = False
    
    '�ɼ�����������
    Set cbrToolBar = Me.cbrMain.Add("������ʷ", xtpBarLeft)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_Mode_Orig, "ԭʼ״̬")
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_Mode_Clear, "����״̬")
        cbrControl.Checked = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "����Ԥ��")
        cbrControl.IconId = 102
        cbrControl.Style = xtpButtonIconAndCaption
        cbrControl.BeginGroup = True
        
        '������ʷ����Ĳ˵���ֻ������ʷ�����ʱ�򣬲���������˵�
        strSQL = "Select ����ID,ִ�п���ID From ����ҽ����¼  Where Id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "", mlngAdviceID)
        If rsTemp.EOF = False Then
            mlngPatientId = nvl(rsTemp!����ID, 0)
            mlngCur����ID = nvl(rsTemp!ִ�п���ID, 0)
            
            strSQL = "Select c.Id As ҽ��id,c.����ʱ�� As ����ʱ��,c.ҽ������,b.����Id  From Ӱ�����¼ a,����ҽ������ b,����ҽ����¼ c" _
                    & " Where a.ҽ��id = c.Id And b.ҽ��ID= c.Id And c.����ID=[1] And c.���ID Is Null And c.ִ�п���ID  in " _
                    & " (Select ����ID From ������Ա Where ��ԱID =[2]) Order By ����ʱ�� Asc"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "", mlngPatientId, UserInfo.ID)
            
            If rsTemp.RecordCount > 1 Or (mlngReportID = 0 And rsTemp.RecordCount = 1) Then
                Set cbrControl = .Add(xtpControlPopup, conMenu_PacsReport_History_Times, "������ʷ"): cbrControl.ID = conMenu_PacsReport_History_Times
                
                Do Until rsTemp.EOF
                   Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_PacsReport_History_Times + rsTemp.AbsolutePosition, "��" & rsTemp.AbsolutePosition & "��(" & Format(rsTemp!����ʱ��, "yyyy-mm-dd") & ") " & rsTemp!ҽ������)
                   cbrPopControl.DescriptionText = rsTemp!ҽ��ID & "-" & rsTemp!����Id
                   rsTemp.MoveNext
                Loop
'                '�����ǰ���ڱ༭�ı��滹û�д����������ӵ�ǰ����Ĳ˵�
'                If mlngReportID = 0 Then
'                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_PacsReport_History_Times + rsTemp.RecordCount + 1, "��ǰ����")
'                   cbrPopControl.DescriptionText = mlngAdviceID & "-0"
'                End If
                mintReportCount = rsTemp.RecordCount
            End If
        End If
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
        cbrControl.Style = xtpButtonIconAndCaption
      
    End With
    cbrToolBar.Position = xtpBarTop
End Sub

Public Sub ShowModeOrig(lngReportID As Long, lngAdviceID As Long)
    
    rtxtReport.Text = ""
    Call ShowTitle(lngReportID, lngAdviceID)
    Call ShowReportText(lngReportID, "�������")
    Call ShowReportText(lngReportID, "������")
    Call ShowReportText(lngReportID, "����")
    
    rtxtReport.SelStart = 0
    rtxtReport.SelLength = 0
End Sub

Private Sub ShowReportText(lngReportID As Long, strType As String)
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim lngStart As Long
    Dim strText As String
    Dim strSeparator2 As String
    Dim strSeparator1 As String
    
    strSeparator1 = vbCrLf & mstrOffset & "-------" & vbCrLf
    strSeparator2 = vbCrLf ' & mstrOffset & "------------" & vbCrLf
    
    
    '��ȡ���������
    strSQL = "Select a.�����ı� As ����, b.��������, b.�����ı� As ����,b.��ʼ�� as �汾 From ���Ӳ������� a,���Ӳ������� b " & _
             " Where a.�ļ�id = [1] And a.�������� = 3 And a.Id = b.��ID And b.�������� = 2 and a.�����ı� =[2] order by b.��ʼ��  "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngReportID, strType)
    
    If rsTemp.EOF = False Then
        lngStart = Len(rtxtReport.Text)
        Select Case strType
            Case "�������"
                strText = strSeparator2 & mstrOffset & pReport_CheckViewName & strSeparator2
            Case "������"
                strText = vbCrLf & strSeparator2 & mstrOffset & pReport_ResultName & strSeparator2
            Case "����"
                strText = vbCrLf & strSeparator2 & mstrOffset & pReport_AdviceName & strSeparator2
        End Select
        
        rtxtReport.SelStart = lngStart
        rtxtReport.SelLength = 0
        rtxtReport.SelText = strText
        
        rtxtReport.SelStart = lngStart
        rtxtReport.SelLength = Len(strText)
        rtxtReport.SelFontName = "����"
        rtxtReport.SelFontSize = 14
        rtxtReport.SelColor = vbBlue
        rtxtReport.SelBold = True
    End If
    
    While Not rsTemp.EOF
        lngStart = Len(rtxtReport.Text)
        strText = strSeparator1 & mstrOffset & "�� " & nvl(rsTemp!�汾) & " �棺" & strSeparator1 & mstrOffset & nvl(rsTemp!����) & vbCrLf
        rtxtReport.SelStart = lngStart
        rtxtReport.SelLength = 0
        rtxtReport.SelText = strText
        
        rtxtReport.SelStart = lngStart
        rtxtReport.SelLength = Len(strText)
        rtxtReport.SelFontName = "����"
        rtxtReport.SelFontSize = 14
        rtxtReport.SelColor = vbBlack
        rtxtReport.SelBold = False
        
        rsTemp.MoveNext
    Wend
End Sub

Public Sub ShowModeClear(lngReportID As Long, lngAdviceID As Long)
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim lngStart As Long
    Dim strText As String
    Dim strTitle As String
    Dim strSeparator2 As String
    Dim blnShow As Boolean
    
    strSeparator2 = vbCrLf 'vbCrLf & mstrOffset & "------------" & vbCrLf
    rtxtReport.Text = ""
    
    Call ShowTitle(lngReportID, lngAdviceID)
    
    '��ȡ���������
    strSQL = "Select a.�����ı� As ����, b.��������, b.�����ı� As ����,b.��ʼ�� as �汾 From ���Ӳ������� a,���Ӳ������� b " & _
             " Where a.�ļ�id = [1] And a.�������� = 3 And a.Id = b.��ID And b.�������� = 2 and b.��ֹ��=0  "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngReportID)
    
    While Not rsTemp.EOF
        blnShow = False
        Select Case rsTemp!����
            Case "�������"
                strTitle = strSeparator2 & mstrOffset & pReport_CheckViewName & strSeparator2
                strText = vbCrLf & mstrOffset & nvl(rsTemp!����) & vbCrLf
                blnShow = True
            Case "������"
                strTitle = strSeparator2 & mstrOffset & pReport_ResultName & strSeparator2
                strText = vbCrLf & mstrOffset & nvl(rsTemp!����) & vbCrLf
                blnShow = True
            Case "����"
                strTitle = strSeparator2 & mstrOffset & pReport_AdviceName & strSeparator2
                strText = vbCrLf & mstrOffset & nvl(rsTemp!����) & vbCrLf
                blnShow = True
        End Select
        
        If blnShow = True Then
            lngStart = Len(rtxtReport.Text)
            rtxtReport.SelStart = lngStart
            rtxtReport.SelLength = 0
            rtxtReport.SelText = strTitle
            
            rtxtReport.SelStart = lngStart
            rtxtReport.SelLength = Len(strTitle)
            rtxtReport.SelFontName = "����"
            rtxtReport.SelFontSize = 14
            rtxtReport.SelColor = vbBlue
            rtxtReport.SelBold = True
            
            lngStart = Len(rtxtReport.Text)
            rtxtReport.SelStart = lngStart
            rtxtReport.SelLength = 0
            rtxtReport.SelText = strText
            
            rtxtReport.SelStart = lngStart
            rtxtReport.SelLength = Len(strText)
            rtxtReport.SelFontName = "����"
            rtxtReport.SelFontSize = 14
            rtxtReport.SelColor = vbBlack
            rtxtReport.SelBold = False
        End If
            
        rsTemp.MoveNext
    Wend
    
    rtxtReport.SelStart = 0
    rtxtReport.SelLength = 0
End Sub
