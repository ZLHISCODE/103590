VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmReportView 
   BorderStyle     =   0  'None
   Caption         =   "��������"
   ClientHeight    =   6135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picSigns 
      Height          =   1455
      Left            =   1680
      ScaleHeight     =   1395
      ScaleWidth      =   3075
      TabIndex        =   3
      Top             =   4560
      Width           =   3135
      Begin VB.TextBox txtReview 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   600
         TabIndex        =   11
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtSigns 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   350
         Left            =   690
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   0
         Width           =   2415
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDD6C6&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label lblReview 
         Caption         =   "��ã�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   12
         Top             =   360
         Width           =   650
      End
      Begin VB.Label lblSign 
         Caption         =   "ǩ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   9
         Top             =   30
         Width           =   650
      End
   End
   Begin VB.PictureBox picAdvice 
      Height          =   1215
      Left            =   3960
      ScaleHeight     =   1155
      ScaleWidth      =   2955
      TabIndex        =   2
      Top             =   3000
      Width           =   3015
      Begin RichTextLib.RichTextBox rTxtAdvice 
         Height          =   855
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1508
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmReportView.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox picResult 
      Height          =   1215
      Left            =   600
      ScaleHeight     =   1155
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   2400
      Width           =   3015
      Begin RichTextLib.RichTextBox rtxtResult 
         Height          =   975
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1720
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmReportView.frx":009D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox picCheckView 
      Height          =   2175
      Left            =   2520
      ScaleHeight     =   2115
      ScaleWidth      =   4275
      TabIndex        =   0
      Top             =   360
      Width           =   4335
      Begin RichTextLib.RichTextBox rtxtCheckView 
         Height          =   1935
         Left            =   840
         TabIndex        =   4
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   3413
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmReportView.frx":013A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblFormat 
         Caption         =   "�±���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   3855
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   600
      Top             =   480
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmReportView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'Private mlngAdviceID As Long    'ҽ��ID
'Private mlngSendNo As Long      '���ͺ�
Private mblnSingleWindow As Boolean     '�Ƿ�ʹ�ö���������ʾ����༭����True-����������ʾ��False-Ƕ��ʽ��ʾ
Private mReportID As Long         '�����ļ�id
Private mFileID As Long           '����ģ��ID
Private mlngCY1 As Long                 '��������ĸ߶�
Private mlngCY2 As Long                 '�������ĸ߶�
Private mlngCY3 As Long                 '����ĸ߶�
Private mlngCY4 As Long                 'ǩ���ĸ߶�
Private mblnCheckModify As Boolean      '�Ƿ��������ݱ仯��¼
Private mblnEdiatble As Boolean         '�Ƿ���Ա༭����
Private mstrModifyEdit As String        '��ǰ�����Ƿ����޶�״̬���������޶������û��ǩ������¼�����˵��������ձ�ʾ�����������
Private mblnShowWord As Boolean         '��ʾ�ʾ�ʾ����True--��ʾ�ʾ�ʾ����False--˫���������ʾ�ʾ�ʾ��
Private mblnMoved As Boolean            '�Ƿ�ת��

Public pModified As Boolean          '��¼��ǰ�����Ƿ��иı�
Private mingFlag As Integer          'Ϊ1ʱ˵���Ѿ�ִ�й����������GetFocue����

'��������¼�
Public Event CheckViewClick(ByVal strContext As String)
Public Event ResultClick(ByVal strContext As String)
Public Event AdviceClick(ByVal strContext As String)
Public Event ShowWord(intReportViewType As Integer, strContext As String)

Public Sub zlRefreshLblFormat(strFormatInfo As String)
    lblFormat.Caption = strFormatInfo
End Sub

Public Sub zlRefresh(ReportID As Long, blnSingleWindow As Boolean, FileID As Long, _
    blnDeptChanged As Boolean, blnEditable As Boolean, strModifyEdit As String, _
    strInfo As String, blnShowWord As Boolean, strFormatInfo As String, ByVal blnMoved As Boolean)
'����˵����
'           blnEditable----��ǰ�Ƿ���Ա༭,True--�ɱ༭��False--���ɱ༭
'           strModifyEdit----��ǰ�����Ƿ����޶�״̬���������޶������û��ǩ������¼�����˵��������ձ�ʾ�����������
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer

    mReportID = ReportID
    mFileID = FileID
    mblnEdiatble = blnEditable
    mstrModifyEdit = strModifyEdit
    mblnShowWord = blnShowWord
    mblnMoved = blnMoved
    mingFlag = 0
    lblFormat.Caption = strFormatInfo
    
    rtxtCheckView.Text = ""
    rtxtResult.Text = ""
    rTxtAdvice.Text = ""
    rtxtCheckView.Tag = ""
    rtxtResult.Tag = ""
    rTxtAdvice.Tag = ""
    
    txtInfo.Visible = blnSingleWindow
    If blnSingleWindow Then txtInfo.Text = strInfo
    
    mblnCheckModify = False         '�ر����ݱ仯��¼
    pModified = False

    If mblnSingleWindow <> blnSingleWindow Then
        mblnSingleWindow = blnSingleWindow
        Call InitLoaclParas     '��ȡ��������
        Call InitFaceScheme     '��ʼ���沼��
    End If
    
    If blnDeptChanged = True Then
        For i = 1 To dkpMain.PanesCount
            Select Case dkpMain.Panes(i).Tag
            Case 0
                dkpMain.Panes(i).Title = pReport_CheckViewName
            Case 1
                dkpMain.Panes(i).Title = pReport_ResultName
            Case 2
                dkpMain.Panes(i).Title = pReport_AdviceName
            End Select
        Next i
    End If
    
    '���ݲ����ļ�ID����ʼ�������ı��༭��
    '���ұ��浥ģ��ID
    If mReportID = 0 Then       '����Ϊ�գ���Ҫ�������棬�ӱ���ģ������ȡ��Ϣ
        strSql = "Select a.�����ı� As ����, b.��������, b.�����ı� As ���� " & _
                 " From �����ļ��ṹ a, �����ļ��ṹ b" & _
                 " Where a.�ļ�id = [1] And a.�������� = 3 And a.Id = b.��id And b.�������� = 2 "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mFileID)
    Else
        strSql = "Select a.�����ı� As ����, b.��������, b.�����ı� As ���� From ���Ӳ������� a,���Ӳ������� b " & _
                 " Where a.�ļ�id = [1] And a.�������� = 3 And a.Id = b.��ID And b.�������� = 2 And b.��ֹ�� = 0"
        If mblnMoved = True Then
            strSql = Replace(strSql, "���Ӳ�������", "H���Ӳ�������")
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mReportID)
    End If
    While rsTemp.EOF = False
        Select Case Nvl(rsTemp!����)
            Case "�������"
                rtxtCheckView.Tag = Nvl(rsTemp!��������)
                zlWriteReport Nvl(rsTemp!����), 0
            Case "������"
                rtxtResult.Tag = Nvl(rsTemp!��������)
                zlWriteReport Nvl(rsTemp!����), 1
            Case "����"
                rTxtAdvice.Tag = Nvl(rsTemp!��������)
                zlWriteReport Nvl(rsTemp!����), 2
        End Select
        rsTemp.MoveNext
    Wend
    
    If mblnSingleWindow Then
        RaiseEvent CheckViewClick("")
    Else
        On Error GoTo errH
        If rtxtCheckView.Visible Then rtxtCheckView.SetFocus
errH:
    End If
    
    '���ý���ؼ��Ƿ���Ա༭
    rtxtCheckView.Locked = Not mblnEdiatble
    rtxtResult.Locked = Not mblnEdiatble
    rTxtAdvice.Locked = Not mblnEdiatble
    
    rtxtCheckView.BackColor = IIf(rtxtCheckView.Locked, &H8000000F, &H80000005)
    rtxtResult.BackColor = IIf(rtxtResult.Locked, &H8000000F, &H80000005)
    rTxtAdvice.BackColor = IIf(rTxtAdvice.Locked, &H8000000F, &H80000005)
    
    rtxtCheckView.ToolTipText = IIf(mblnEdiatble = False And mstrModifyEdit <> "", "�������Ѿ���" & mstrModifyEdit & "�����޶��������Ƿ�Ҫ�����޶�����Ҫ�޶���˫����", "")
    rtxtResult.ToolTipText = rtxtCheckView.ToolTipText
    rTxtAdvice.ToolTipText = rtxtCheckView.ToolTipText
    
'    picCheckView.Enabled = mblnEdiatble
'    picResult.Enabled = mblnEdiatble
'    picAdvice.Enabled = mblnEdiatble
    
    mblnCheckModify = True      '����װ����ϣ��������ݱ仯��¼
End Sub

Public Sub zlWriteReport(strText As String, intType As Integer)
    'intType---0 ���������1 ��������2 ����
    Dim rText As RichTextBox
    Dim lngCount As Long
    Dim lngSelStart As Long
    Dim lngPosStart As Long
    Dim lngPosEnd As Long
    
    On Error GoTo err
    
    If intType = 0 Then
        Set rText = rtxtCheckView
    ElseIf intType = 1 Then
        Set rText = rtxtResult
    ElseIf intType = 2 Then
        Set rText = rTxtAdvice
    End If
    
    lngSelStart = rText.SelStart
    rText.SelLength = 0
    rText.SelText = strText
    '������ɫ
    rText.SelStart = lngSelStart
    rText.SelLength = Len(strText)
    rText.SelColor = vbBlack
    
    On Error Resume Next
    'rText.Tag �ǵ��Ӳ�����ʽ�Ķ������ԣ��á�|���ָ����ܹ�26��Ԫ��
    rText.SelStart = 0
    rText.SelLength = Len(rText.Text)
    rText.SelFontName = Split(rText.Tag, "|")(15)     '  rText.SelFontName
    rText.SelFontSize = Split(rText.Tag, "|")(16)     ' rText.SelFontSize
    rText.SelBold = Split(rText.Tag, "|")(17)     'rText.SelBold
    rText.SelItalic = Split(rText.Tag, "|")(18)   'rText.SelItalic
    On Error GoTo 0
    
    '������ǰ��������֣��Ƿ���Ҫ�أ������������ɫ��ʾ����
    '�Ȳ��ѡҪ��
    For lngCount = 1 To Len(strText)
        lngPosStart = InStr(lngCount, strText, "{{")
        lngPosEnd = InStr(lngCount, strText, "}}")
        If lngPosStart <> 0 And lngPosEnd <> 0 And lngPosEnd > lngPosStart Then
            '���ҵ�Ҫ�أ����Ҫ������ɫ��ʾ
            rText.SelStart = lngSelStart + lngPosStart - 1
            rText.SelLength = lngPosEnd - lngPosStart + 2
            rText.SelColor = vbBlue
            lngCount = lngPosEnd
        Else
            Exit For
        End If
    Next lngCount
    
    '�ٲ鵥ѡҪ��
    For lngCount = 1 To Len(strText)
        lngPosStart = InStr(lngCount, strText, "{<")
        lngPosEnd = InStr(lngCount, strText, ">}")
        If lngPosStart <> 0 And lngPosEnd <> 0 And lngPosEnd > lngPosStart Then
            '���ҵ�Ҫ�أ����Ҫ������ɫ��ʾ
            rText.SelStart = lngSelStart + lngPosStart - 1
            rText.SelLength = lngPosEnd - lngPosStart + 2
            rText.SelColor = vbBlue
            lngCount = lngPosEnd
        Else
            Exit For
        End If
    Next lngCount
    
    rText.SelStart = lngSelStart + Len(strText)
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    Dim strContext As String
    
    Select Case Pane.Tag
        Case 0
            strContext = rtxtCheckView.Text
        Case 1
            strContext = rtxtResult.Text
        Case 2
            strContext = rTxtAdvice.Text
    End Select
    
    If Action = PaneActionFloating And mblnShowWord = False Then
'        frmReportWord.Show 1, Me
    
        '�����¼�����ʾ�ʾ�ʾ������
        RaiseEvent ShowWord(Pane.Tag, strContext)
    End If
    Cancel = True
End Sub

Private Sub Form_Load()
    mingFlag = 0
    pModified = False
    mblnSingleWindow = False    'Ĭ������ΪǶ��ʽ����
    
    Call InitLoaclParas     '��ȡ��������
    Call InitFaceScheme     '��ʼ���沼��
End Sub

Private Sub InitLoaclParas()
    Dim strRegPath As String
    
    '��ȡ��������������������򣬽������� ��ǩ������ĸ߶�
    If mblnSingleWindow = True Then
        strRegPath = "����ģ��\" & App.ProductName & "\frmReportView\SingleWindow"
    Else
        strRegPath = "����ģ��\" & App.ProductName & "\frmReportView"
    End If
    mlngCY1 = GetSetting("ZLSOFT", strRegPath, "CY1", 500)
    mlngCY2 = GetSetting("ZLSOFT", strRegPath, "CY2", 200)
    mlngCY3 = GetSetting("ZLSOFT", strRegPath, "CY3", 100)
    mlngCY4 = GetSetting("ZLSOFT", strRegPath, "CY4", 100)
End Sub

Private Sub InitFaceScheme()
    '��ʼ���沼��
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane, pane4 As Pane
    With Me.dkpMain
        .CloseAll
        .Options.HideClient = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    Set Pane1 = dkpMain.CreatePane(1, 0, mlngCY1, DockTopOf, Nothing)
    Pane1.Title = pReport_CheckViewName
    Pane1.Handle = picCheckView.hWnd
    Pane1.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Pane1.Tag = 0
    
    Set Pane2 = dkpMain.CreatePane(2, 0, mlngCY2, DockBottomOf, Pane1)
    Pane2.Title = pReport_ResultName
    Pane2.Handle = picResult.hWnd
    Pane2.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Pane2.Tag = 1
    
    Set Pane3 = dkpMain.CreatePane(3, 0, mlngCY3, DockBottomOf, Pane2)
    Pane3.Title = pReport_AdviceName
    Pane3.Handle = picAdvice.hWnd
    Pane3.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Pane3.Tag = 2
    
    Set pane4 = dkpMain.CreatePane(4, 0, mlngCY4, DockBottomOf, Pane3)
    pane4.Title = "ǩ��"
    pane4.Handle = picSigns.hWnd
    pane4.Options = PaneNoCaption Or PaneNoCloseable
    pane4.Tag = 3
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strRegPath As String
    
    If mblnSingleWindow = True Then
        strRegPath = "����ģ��\" & App.ProductName & "\frmReportView\SingleWindow"
    Else
        strRegPath = "����ģ��\" & App.ProductName & "\frmReportView"
    End If
    '�����������������������򣬽��������ǩ������ĸ߶�
    '285��Pane�ı���߶ȣ�ʹ���˱��⣬����Ҫ�ӻ�����߶�
    SaveSetting "ZLSOFT", strRegPath, "CY1", picCheckView.Height + 285
    SaveSetting "ZLSOFT", strRegPath, "CY2", picResult.Height + 285
    SaveSetting "ZLSOFT", strRegPath, "CY3", picAdvice.Height + 285
    SaveSetting "ZLSOFT", strRegPath, "CY4", picSigns.Height
    
    If mblnSingleWindow = True Then
        strRegPath = "����ģ��\" & App.ProductName & "\frmReport\SingleWindow"
    Else
        strRegPath = "����ģ��\" & App.ProductName & "\frmReport"
    End If
    SaveSetting "ZLSOFT", strRegPath, "CX2", Me.Width
    SaveSetting "ZLSOFT", strRegPath, "CY21", Me.Height
End Sub

Private Sub picAdvice_Resize()
On Error Resume Next

    rTxtAdvice.Left = 20
    rTxtAdvice.Top = 20
    If picAdvice.Height > 50 And picAdvice.Width > 50 Then
        rTxtAdvice.Width = Abs(picAdvice.Width - 100)
        rTxtAdvice.Height = Abs(picAdvice.Height - 100)
    End If
End Sub

Private Sub picCheckView_Resize()
On Error Resume Next

    lblFormat.Left = 10
    lblFormat.Top = 10
    lblFormat.Width = picCheckView.Width
    lblFormat.Height = 400
    
    rtxtCheckView.Left = 20
    rtxtCheckView.Top = lblFormat.Height
    If picCheckView.Width > 50 And picCheckView.Height > 50 Then
        rtxtCheckView.Width = Abs(picCheckView.Width - 100)
        rtxtCheckView.Height = Abs(picCheckView.Height - 100 - lblFormat.Height)
    End If
End Sub

Private Sub picResult_Resize()
On Error Resume Next

    rtxtResult.Left = 20
    rtxtResult.Top = 20
    If picResult.Width > 50 And picResult.Height > 50 Then
        rtxtResult.Width = Abs(picResult.Width - 100)
        rtxtResult.Height = Abs(picResult.Height - 100)
    End If
End Sub

Private Sub picSigns_Resize()
On Error Resume Next

    lblSign.Left = 0
    txtSigns.Left = lblSign.Width
    txtSigns.Top = 30
    txtSigns.Width = Abs(picSigns.ScaleWidth - lblSign.Width - 50)
    
    lblReview.Left = 0
    txtReview.Left = lblSign.Width
    txtReview.Top = 360
    txtReview.Width = txtSigns.Width
    
    txtInfo.Left = 10
    txtInfo.Top = txtReview.Top + txtReview.Height + 10
    txtInfo.Width = Abs(picSigns.ScaleWidth - 50)
    txtInfo.Height = Abs(picSigns.ScaleHeight - txtSigns.Height - txtReview.Height - 50)
End Sub

Private Sub rTxtAdvice_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub rTxtAdvice_DblClick()
    If mblnEdiatble = False And mstrModifyEdit <> "" Then
        rTxtAdvice.Locked = False
        rTxtAdvice.ToolTipText = ""
    Else
        Call richTextBoxShowElements(rTxtAdvice)
    End If
End Sub

Private Sub rTxtAdvice_GotFocus()
On Error GoTo err
    If gblnIsStudyChage Then
        If rtxtCheckView.Visible Then rtxtCheckView.SetFocus '�л����󣬶�λ����������ı��򣬾��������81704
        gblnIsStudyChage = False
        Exit Sub
    End If
    mingFlag = 0
    
    Call zlCommFun.OpenIme(True)
    RaiseEvent AdviceClick(rTxtAdvice.Text)
err:
End Sub
 

Private Sub rtxtCheckView_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub rtxtCheckView_DblClick()
    If mblnEdiatble = False And mstrModifyEdit <> "" Then
        rtxtCheckView.Locked = False
        rtxtCheckView.ToolTipText = ""
    Else
        Call richTextBoxShowElements(rtxtCheckView)
    End If
End Sub

Private Sub rtxtCheckView_GotFocus()
    If mingFlag = 1 Then Exit Sub
    
    mingFlag = 1
    Call zlCommFun.OpenIme(True)
    RaiseEvent CheckViewClick(rtxtCheckView.Text)
End Sub
 
Private Sub rtxtResult_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub rtxtResult_DblClick()
    If mblnEdiatble = False And mstrModifyEdit <> "" Then
        rtxtResult.Locked = False
        rtxtResult.ToolTipText = ""
    Else
        Call richTextBoxShowElements(rtxtResult)
    End If
End Sub

Private Sub rtxtResult_GotFocus()
On Error GoTo err
    If gblnIsStudyChage Then
        If rtxtCheckView.Visible Then rtxtCheckView.SetFocus '�л����󣬶�λ����������ı��򣬾��������81704
        gblnIsStudyChage = False
        Exit Sub
    End If
    mingFlag = 0
    
    Call zlCommFun.OpenIme(True)
    RaiseEvent ResultClick(rtxtResult.Text)
err:
End Sub
 

Private Sub txtReview_Change()
    If mblnCheckModify = True Then
        pModified = True
    End If
End Sub

Private Sub txtReview_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub
 
