VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmReportWordEdit 
   Caption         =   "����ʾ�༭"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9720
   Icon            =   "frmReportWordEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   9720
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picClientArea 
      Height          =   5175
      Left            =   120
      ScaleHeight     =   5115
      ScaleWidth      =   9315
      TabIndex        =   2
      Top             =   0
      Width           =   9375
      Begin VB.PictureBox picAdvice 
         AutoSize        =   -1  'True
         Height          =   1575
         Left            =   3720
         ScaleHeight     =   1515
         ScaleWidth      =   5595
         TabIndex        =   13
         Top             =   3480
         Width           =   5655
         Begin VB.VScrollBar vscroWordH3 
            Height          =   1215
            Left            =   5280
            Max             =   500
            TabIndex        =   17
            Top             =   240
            Value           =   200
            Width           =   250
         End
         Begin VB.PictureBox picContainer3 
            BorderStyle     =   0  'None
            Height          =   1095
            Left            =   120
            ScaleHeight     =   1095
            ScaleWidth      =   5175
            TabIndex        =   14
            Top             =   240
            Width           =   5175
            Begin VB.CheckBox chkSelect3 
               DownPicture     =   "frmReportWordEdit.frx":0CCA
               Height          =   400
               Index           =   0
               Left            =   480
               Picture         =   "frmReportWordEdit.frx":1C3C
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   120
               Visible         =   0   'False
               Width           =   400
            End
            Begin RichTextLib.RichTextBox rtxtWord3 
               Height          =   495
               Index           =   0
               Left            =   960
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   120
               Visible         =   0   'False
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   873
               _Version        =   393217
               Enabled         =   -1  'True
               ScrollBars      =   2
               Appearance      =   0
               AutoVerbMenu    =   -1  'True
               TextRTF         =   $"frmReportWordEdit.frx":2BAE
            End
         End
      End
      Begin VB.PictureBox picResult 
         AutoSize        =   -1  'True
         Height          =   1695
         Left            =   120
         ScaleHeight     =   1635
         ScaleWidth      =   5595
         TabIndex        =   8
         Top             =   1800
         Width           =   5655
         Begin VB.PictureBox picContainer2 
            BorderStyle     =   0  'None
            Height          =   975
            Left            =   120
            ScaleHeight     =   975
            ScaleWidth      =   5175
            TabIndex        =   10
            Top             =   240
            Width           =   5175
            Begin VB.CheckBox chkSelect2 
               DownPicture     =   "frmReportWordEdit.frx":2C4B
               Height          =   400
               Index           =   0
               Left            =   480
               Picture         =   "frmReportWordEdit.frx":3BBD
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   120
               Visible         =   0   'False
               Width           =   400
            End
            Begin RichTextLib.RichTextBox rtxtWord2 
               Height          =   375
               Index           =   0
               Left            =   960
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   120
               Visible         =   0   'False
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   661
               _Version        =   393217
               Enabled         =   -1  'True
               ScrollBars      =   2
               Appearance      =   0
               AutoVerbMenu    =   -1  'True
               TextRTF         =   $"frmReportWordEdit.frx":4B2F
            End
         End
         Begin VB.VScrollBar vscroWordH2 
            Height          =   1215
            Left            =   5280
            Max             =   500
            TabIndex        =   9
            Top             =   240
            Value           =   200
            Width           =   250
         End
      End
      Begin VB.PictureBox picCheckView 
         AutoSize        =   -1  'True
         Height          =   1575
         Left            =   120
         ScaleHeight     =   1515
         ScaleWidth      =   9075
         TabIndex        =   3
         Top             =   120
         Width           =   9135
         Begin VB.VScrollBar vscroWordH1 
            Height          =   1215
            Left            =   5280
            Max             =   500
            TabIndex        =   7
            Top             =   240
            Value           =   200
            Width           =   250
         End
         Begin VB.PictureBox picContainer1 
            BorderStyle     =   0  'None
            Height          =   975
            Left            =   120
            ScaleHeight     =   975
            ScaleWidth      =   5175
            TabIndex        =   4
            Top             =   240
            Width           =   5175
            Begin VB.CheckBox chkSelect1 
               DownPicture     =   "frmReportWordEdit.frx":4BCC
               Height          =   400
               Index           =   0
               Left            =   480
               Picture         =   "frmReportWordEdit.frx":5B3E
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   120
               Visible         =   0   'False
               Width           =   400
            End
            Begin RichTextLib.RichTextBox rtxtWord1 
               Height          =   375
               Index           =   0
               Left            =   960
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   120
               Visible         =   0   'False
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   661
               _Version        =   393217
               Enabled         =   -1  'True
               ScrollBars      =   2
               Appearance      =   0
               AutoVerbMenu    =   -1  'True
               TextRTF         =   $"frmReportWordEdit.frx":6AB0
            End
         End
      End
      Begin XtremeDockingPane.DockingPane dkpMain 
         Left            =   240
         Top             =   3840
         _Version        =   589884
         _ExtentX        =   450
         _ExtentY        =   423
         _StockProps     =   0
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      Height          =   400
      Left            =   6360
      TabIndex        =   1
      Top             =   6000
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Height          =   400
      Left            =   2520
      TabIndex        =   0
      Top             =   6000
      Width           =   1100
   End
End
Attribute VB_Name = "frmReportWordEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private miWordScale As Integer
Private mlngWordID As Long
Private mstrCheckView As String
Private mstrResult As String
Private mstrAdvice As String
Private miType As Integer

Public Sub zlShowMe(lngWordID As Long, frmParent As Object, iType As Integer, ByRef strCheckView As String, ByRef strResult As String, ByRef strAdvice As String)
    mlngWordID = lngWordID
    miType = iType
    frmReportWordEdit.Show 1, frmParent
    strCheckView = mstrCheckView
    strResult = mstrResult
    strAdvice = mstrAdvice
End Sub

Private Sub FillWords(lngWordID As Long)
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim blnNextLine As Boolean      '�Ƿ���һ�У�������������ؼ�
    Dim iFieldCount As Integer
    Dim iType As Integer        '0-������1-��ϣ�2-���
    Dim blnStartSegment As Boolean      '��ʼһ������
    
    '���ԭ�пؼ�
    Call ClearWordShow
    blnNextLine = True
    miWordScale = 0
    
    strSQL = "Select �ʾ�id,���д���,��������,�����ı�,����Ҫ��ID,�滻��,Ҫ������,Ҫ������,Ҫ�س���,Ҫ��С��," & _
             " Ҫ�ص�λ,Ҫ�ر�ʾ,Ҫ��ֵ��,������̬ From �����ʾ���� Where �ʾ�ID=[1] order by ���д��� "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngWordID)
    
    blnStartSegment = False
    
    On Error GoTo ErrHandle
    '����ÿһ�У���ʾ
    While rsTemp.EOF = False
        
        If blnNextLine = True Then
            blnNextLine = False
            
            '�ȶ�ȡ�����ı����жϵ�ǰ���ݵ�����
            If Left(Nvl(rsTemp!�����ı�), 6) = "<<����>>" Then
                iType = 0
            ElseIf Left(Nvl(rsTemp!�����ı�), 6) = "<<���>>" Then
                iType = 1
            ElseIf Left(Nvl(rsTemp!�����ı�), 6) = "<<����>>" Then
                iType = 2
            Else
                iType = miType
            End If
            
            '������Ӧ���͵Ŀؼ�
            If iType = 0 Then           '������������Ŀؼ�
                iFieldCount = rtxtWord1.Count
                '������ť���ı���
                Load rtxtWord1(iFieldCount)
                rtxtWord1(iFieldCount).Visible = True
                Load chkSelect1(iFieldCount)
                chkSelect1(iFieldCount).Visible = True
                
                '�ڷ�λ��
                If iFieldCount = 1 Then
                    chkSelect1(iFieldCount).Top = 5
                Else
                    chkSelect1(iFieldCount).Top = rtxtWord1(iFieldCount - 1).Top + rtxtWord1(iFieldCount - 1).Height + 5
                End If
                chkSelect1(iFieldCount).Left = 150
                rtxtWord1(iFieldCount).Left = chkSelect1(iFieldCount).Left + chkSelect1(iFieldCount).Width + 150
                rtxtWord1(iFieldCount).Top = chkSelect1(iFieldCount).Top
                rtxtWord1(iFieldCount).Width = picContainer1.Width - rtxtWord1(iFieldCount).Left - 60
                rtxtWord1(iFieldCount).Height = 400
            ElseIf iType = 1 Then       '�����������Ŀؼ�
                iFieldCount = rtxtWord2.Count
                '������ť���ı���
                Load rtxtWord2(iFieldCount)
                rtxtWord2(iFieldCount).Visible = True
                Load chkSelect2(iFieldCount)
                chkSelect2(iFieldCount).Visible = True
                
                '�ڷ�λ��
                If iFieldCount = 1 Then
                    chkSelect2(iFieldCount).Top = 5
                Else
                    chkSelect2(iFieldCount).Top = rtxtWord2(iFieldCount - 1).Top + rtxtWord2(iFieldCount - 1).Height + 5
                End If
                chkSelect2(iFieldCount).Left = 150
                rtxtWord2(iFieldCount).Left = chkSelect2(iFieldCount).Left + chkSelect2(iFieldCount).Width + 150
                rtxtWord2(iFieldCount).Top = chkSelect2(iFieldCount).Top
                rtxtWord2(iFieldCount).Width = picContainer2.Width - rtxtWord2(iFieldCount).Left - 60
                rtxtWord2(iFieldCount).Height = 400
            ElseIf iType = 2 Then       '��������Ŀؼ�
                iFieldCount = rtxtWord3.Count
                '������ť���ı���
                Load rtxtWord3(iFieldCount)
                rtxtWord3(iFieldCount).Visible = True
                Load chkSelect3(iFieldCount)
                chkSelect3(iFieldCount).Visible = True
                
                '�ڷ�λ��
                If iFieldCount = 1 Then
                    chkSelect3(iFieldCount).Top = 5
                Else
                    chkSelect3(iFieldCount).Top = rtxtWord3(iFieldCount - 1).Top + rtxtWord3(iFieldCount - 1).Height + 5
                End If
                chkSelect3(iFieldCount).Left = 150
                rtxtWord3(iFieldCount).Left = chkSelect3(iFieldCount).Left + chkSelect3(iFieldCount).Width + 150
                rtxtWord3(iFieldCount).Top = chkSelect3(iFieldCount).Top
                rtxtWord3(iFieldCount).Width = picContainer3.Width - rtxtWord3(iFieldCount).Left - 60
                rtxtWord3(iFieldCount).Height = 400
            End If
        End If
        
        'д��rtxt�ؼ�
        If iType = 0 Then
            WriteIntoRTxt rtxtWord1(iFieldCount), Val(Nvl(rsTemp!��������)), Nvl(rsTemp!�����ı�), Val(Nvl(rsTemp!Ҫ�ر�ʾ)), _
                    Nvl(rsTemp!Ҫ�ص�λ), Nvl(rsTemp!Ҫ��ֵ��), blnStartSegment, blnNextLine
        ElseIf iType = 1 Then
            WriteIntoRTxt rtxtWord2(iFieldCount), Val(Nvl(rsTemp!��������)), Nvl(rsTemp!�����ı�), Val(Nvl(rsTemp!Ҫ�ر�ʾ)), _
                    Nvl(rsTemp!Ҫ�ص�λ), Nvl(rsTemp!Ҫ��ֵ��), blnStartSegment, blnNextLine
        ElseIf iType = 2 Then
            WriteIntoRTxt rtxtWord3(iFieldCount), Val(Nvl(rsTemp!��������)), Nvl(rsTemp!�����ı�), Val(Nvl(rsTemp!Ҫ�ر�ʾ)), _
                    Nvl(rsTemp!Ҫ�ص�λ), Nvl(rsTemp!Ҫ��ֵ��), blnStartSegment, blnNextLine
        End If
        
        rsTemp.MoveNext
    Wend
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub WriteIntoRTxt(ByRef rtxtWord As RichTextBox, int�������� As Integer, str�����ı� As String, _
                intҪ�ر�ʾ As Integer, strҪ�ص�λ As String, strҪ��ֵ�� As String, ByRef blnStartSegment As Boolean, ByRef blnNextLine As Boolean)

    If int�������� = 0 Then     '�������ı���ֱ�Ӽ�������
        If Trim(str�����ı�) <> "" And Trim(str�����ı�) <> vbCrLf Then
            
            '��������
            rtxtWord.SelStart = Len(rtxtWord.Text)
            rtxtWord.SelLength = 0
            rtxtWord.SelColor = vbBlack
            '������ִ�ǰ���б�����дλ�ñ�ʶ��ɾ���ñ�ʶ
            If Left(str�����ı�, 6) = "<<����>>" Or Left(str�����ı�, 6) = "<<���>>" _
                Or Left(str�����ı�, 6) = "<<����>>" Then
                rtxtWord.SelText = Right(str�����ı�, Len(str�����ı�) - 6)
            ElseIf UCase(Left(str�����ı�, 3)) = "<P>" Then
                '�ж��Ƿ�<P>��</P>��Χ��һ�������Ķ�
                If UCase(Right(str�����ı�, 4)) = "</P>" Then
                    rtxtWord.SelText = Mid(str�����ı�, 4, Len(str�����ı�) - 7)
                ElseIf UCase(Right(str�����ı�, 6)) = "</P>" & vbCrLf Then
                    rtxtWord.SelText = Mid(str�����ı�, 4, Len(str�����ı�) - 9)
                Else
                    rtxtWord.SelText = Right(str�����ı�, Len(str�����ı�) - 3)
                End If
                blnStartSegment = True
            ElseIf UCase(Right(str�����ı�, 4)) = "</P>" Then
                rtxtWord.SelText = Left(str�����ı�, Len(str�����ı�) - 4)
            ElseIf UCase(Right(str�����ı�, 6)) = "</P>" & vbCrLf Then
                rtxtWord.SelText = Left(str�����ı�, Len(str�����ı�) - 6)
            Else
                rtxtWord.SelText = str�����ı�
            End If
            
            If blnStartSegment = True Then      '�Ѿ����ö����ǣ�����ҽ�������ı��</P>
                If UCase(Right(str�����ı�, 4)) = "</P>" Or UCase(Right(str�����ı�, 6)) = "</P>" & vbCrLf Then
                    blnNextLine = True
                    blnStartSegment = False
                End If
            Else    '���һس���Ϊ����������
                If Right(str�����ı�, 2) = vbCrLf Then
                    blnNextLine = True
                End If
            End If
        End If
    Else        '��Ҫ�أ���Ҫ����
        If intҪ�ر�ʾ = 0 Then     '�ı�Ҫ�ؽ����ɿա� ��
            rtxtWord.SelStart = Len(rtxtWord.Text)
            rtxtWord.SelLength = 0
            rtxtWord.SelText = "  " & strҪ�ص�λ
            
            rtxtWord.SelStart = Len(rtxtWord.Text) - Len(strҪ�ص�λ)
            rtxtWord.SelLength = Len("  " & strҪ�ص�λ)
            rtxtWord.SelColor = vbBlue
        ElseIf intҪ�ر�ʾ = 1 Then     '����
            'Ŀǰû��ʹ�������ʽ
        ElseIf intҪ�ر�ʾ = 2 Then     '��ѡ
            rtxtWord.SelStart = Len(rtxtWord.Text)
            rtxtWord.SelLength = 0
            rtxtWord.SelText = "{{" & strҪ��ֵ�� & "}}" & strҪ�ص�λ
            
            rtxtWord.SelStart = Len(rtxtWord.Text) - Len("{{" & strҪ��ֵ�� & "}}" & strҪ�ص�λ)
            rtxtWord.SelLength = Len("{{" & strҪ��ֵ�� & "}}" & strҪ�ص�λ)
            rtxtWord.SelColor = vbBlue
        ElseIf intҪ�ر�ʾ = 3 Then     '��ѡ
            rtxtWord.SelStart = Len(rtxtWord.Text)
            rtxtWord.SelLength = 0
            rtxtWord.SelText = "{<" & strҪ��ֵ�� & ">}" & strҪ�ص�λ
            
            rtxtWord.SelStart = Len(rtxtWord.Text) - Len("{<" & strҪ��ֵ�� & ">}" & strҪ�ص�λ)
            rtxtWord.SelLength = Len("{<" & strҪ��ֵ�� & ">}" & strҪ�ص�λ)
            rtxtWord.SelColor = vbBlue
        End If
    End If
    
    '��������
    rtxtWord.SelStart = 0
    rtxtWord.SelLength = Len(rtxtWord.Text)
    rtxtWord.SelFontSize = 14
    rtxtWord.SelLength = 0
    
    ResizeRichTextBox rtxtWord
    If rtxtWord.Index = 1 Then
        miWordScale = rtxtWord.Height / IIf(Len(rtxtWord.Text) = 0, 1, Len(rtxtWord.Text))
    End If
End Sub

Private Sub InitFaceScheme()
    '��ʼ���沼��
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane
    With Me.dkpMain
        .CloseAll
        .Options.HideClient = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    Set Pane1 = dkpMain.CreatePane(1, 0, 300, DockTopOf, Nothing)
    Pane1.Title = pReport_CheckViewName
    Pane1.Handle = picCheckView.Hwnd
    Pane1.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set Pane2 = dkpMain.CreatePane(2, 0, 150, DockBottomOf, Pane1)
    Pane2.Title = pReport_ResultName
    Pane2.Handle = picResult.Hwnd
    Pane2.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set Pane3 = dkpMain.CreatePane(3, 0, 80, DockBottomOf, Pane2)
    Pane3.Title = pReport_AdviceName
    Pane3.Handle = picAdvice.Hwnd
    Pane3.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    '��ѡ�е����ݣ���֯������
    Dim strCheckView As String
    Dim strResult As String
    Dim strAdvice As String
    Dim i As Integer
    
    For i = 1 To chkSelect1.Count - 1
        If chkSelect1(i).Value = 1 Then
            If Right(rtxtWord1(i).Text, 2) = vbCrLf Then
                strCheckView = strCheckView & Left(rtxtWord1(i).Text, Len(rtxtWord1(i).Text) - 2)
            Else
                strCheckView = strCheckView & rtxtWord1(i).Text
            End If
        End If
    Next i
    
    For i = 1 To chkSelect2.Count - 1
        If chkSelect2(i).Value = 1 Then
            If Right(rtxtWord2(i).Text, 2) = vbCrLf Then
                strResult = strResult & Left(rtxtWord2(i).Text, Len(rtxtWord2(i).Text) - 2)
            Else
                strResult = strResult & rtxtWord2(i).Text
            End If
        End If
    Next i
    
    For i = 1 To chkSelect3.Count - 1
        If chkSelect3(i).Value = 1 Then
            If Right(rtxtWord3(i).Text, 2) = vbCrLf Then
                strAdvice = strAdvice & Left(rtxtWord3(i).Text, Len(rtxtWord3(i).Text) - 2)
            Else
                strAdvice = strAdvice & rtxtWord3(i).Text
            End If
        End If
    Next i
    
    mstrCheckView = strCheckView
    mstrResult = strResult
    mstrAdvice = strAdvice
    
    Unload Me
End Sub

Private Sub Form_Load()
    
    mstrCheckView = ""
    mstrResult = ""
    mstrAdvice = ""
    
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitFaceScheme '��ʼ������
    
    Call FillWords(mlngWordID)
End Sub

Private Sub Form_Resize()
    '������ʾ�Ŀͻ�����
    Me.picClientArea.Left = 0
    Me.picClientArea.Top = 0
    Me.picClientArea.Width = Me.ScaleWidth
    Me.picClientArea.Height = Abs(Me.ScaleHeight - 800)
    
    Me.cmdOk.Left = Me.ScaleWidth / 4
    Me.cmdOk.Top = Me.ScaleHeight - 600
    
    Me.cmdCancel.Left = Me.ScaleWidth / 4 * 3 - Me.cmdCancel.Width
    Me.cmdCancel.Top = Me.cmdOk.Top
    
    '�����ʾ�������λ�úͿ��
    picContainer1.Left = 0
    picContainer1.Top = 0
    picContainer1.Width = Abs(picClientArea.Width - vscroWordH1.Width - 60)
    
    picContainer2.Left = 0
    picContainer2.Top = 0
    picContainer2.Width = Abs(picClientArea.Width - vscroWordH1.Width - 60)
    
    picContainer3.Left = 0
    picContainer3.Top = 0
    picContainer3.Width = Abs(picClientArea.Width - vscroWordH1.Width - 60)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '���洰��λ��
    Call SaveWinState(Me, App.ProductName)
End Sub

Public Function ResizeRichTextBox(ByRef rtxtBox As RichTextBox) As Boolean           '�жϴ�ֱ�������Ŀɼ���
    Dim strSegment() As String
    Dim lngWordCount As Long
    Dim i As Integer
    
    lngWordCount = rtxtBox.Width / 350
    If Len(rtxtBox.Text) = 0 Then
        rtxtBox.Height = 400
    Else
        'ÿ����ռ�ÿ��370
        If InStr(rtxtBox.Text, vbCrLf) Then
            strSegment() = Split(rtxtBox.Text, vbCrLf)
            rtxtBox.Height = 0
            For i = 0 To UBound(strSegment)
                rtxtBox.Height = rtxtBox.Height + 400 * (Len(strSegment(i)) / lngWordCount + 1)
            Next i
        Else
            rtxtBox.Height = 400 * (Len(rtxtBox.Text) / lngWordCount + 1)
        End If
    End If
End Function

Public Function ResizeRichTextBox1(ByRef rtxtBox As RichTextBox) As Boolean           '�жϴ�ֱ�������Ŀɼ���
    Dim wndStyle As Long
    Dim i As Integer
    
    i = 0
    rtxtBox.Refresh
    wndStyle = GetWindowLong(rtxtBox.Hwnd, GWL_STYLE)
    
    While (wndStyle And WS_VSCROLL) <> 0 And i < 20
        rtxtBox.Height = rtxtBox.Height + 200
        rtxtBox.Refresh
        If miWordScale <> 0 Then
            '�жϵ�ǰ�߶Ⱥ���������֮��ı����Ƿ���ڵ�һ���ı���ñ�����2��
            If rtxtBox.Height / Len(rtxtBox.Text) > miWordScale * 2 Then
                i = 20
            End If
        End If
        wndStyle = GetWindowLong(rtxtBox.Hwnd, GWL_STYLE)
        i = i + 1
    Wend
End Function

Private Sub ClearWordShow()
    Dim i As Integer
    
    For i = 1 To rtxtWord1.Count - 1
        Unload rtxtWord1(i)
    Next i
    For i = 1 To chkSelect1.Count - 1
        Unload chkSelect1(i)
    Next i
    
    For i = 1 To rtxtWord2.Count - 1
        Unload rtxtWord2(i)
    Next i
    For i = 1 To chkSelect2.Count - 1
        Unload chkSelect2(i)
    Next i
    
    For i = 1 To rtxtWord3.Count - 1
        Unload rtxtWord3(i)
    Next i
    For i = 1 To chkSelect3.Count - 1
        Unload chkSelect3(i)
    Next i
End Sub

Private Sub ResizeWordContainer(picWordShow As PictureBox, vscroWordH As VScrollBar, picWordContainer As PictureBox, lngH As Long)
    
    '������������λ�ú͸߶�
    vscroWordH.Left = picWordShow.Width - vscroWordH.Width
    vscroWordH.Top = 0
    vscroWordH.Height = picWordShow.Height
    
    '�����ʾ�������λ�úͿ��
    picWordContainer.Left = 0
    picWordContainer.Top = 0
    picWordContainer.Width = Abs(picWordShow.Width - vscroWordH.Width)
    
    '�����ʾ������ĸ߶�
    
    If lngH < picWordShow.Height Then
        picWordContainer.Height = picWordShow.Height
        vscroWordH.Enabled = False
    Else
        picWordContainer.Height = lngH
        vscroWordH.Enabled = True
    End If
    
    '���ù������ķ���
    vscroWordH.Max = picWordContainer.Height / 1000
    vscroWordH.Value = 0
End Sub

Private Sub picAdvice_Resize()
    Dim i As Integer
    Dim lngH As Long
    
    '����ÿһ��RichTextBox�Ŀ��
    For i = 1 To rtxtWord3.Count - 1
        rtxtWord3(i).Width = Abs(picContainer3.Width - rtxtWord3(i).Left - 60)
    Next i
    
    '���ڴʾ������ĸ߶�
    For i = 1 To rtxtWord3.Count - 1
        ResizeRichTextBox rtxtWord3(i)
        If i = 1 Then
            rtxtWord3(i).Top = 30
        Else
            rtxtWord3(i).Top = rtxtWord3(i - 1).Top + rtxtWord3(i - 1).Height + 5
        End If
        chkSelect3(i).Top = rtxtWord3(i).Top
    Next
    
    lngH = 0
    If rtxtWord3.Count > 1 Then
        lngH = rtxtWord3(rtxtWord3.Count - 1).Top + rtxtWord3(rtxtWord3.Count - 1).Height + 200
    End If
    
    Call ResizeWordContainer(picAdvice, vscroWordH3, picContainer3, lngH)
End Sub

Private Sub picCheckView_Resize()
    Dim i As Integer
    Dim lngH As Long
    
    '����ÿһ��RichTextBox�Ŀ��
    For i = 1 To rtxtWord1.Count - 1
        rtxtWord1(i).Width = Abs(picContainer1.Width - rtxtWord1(i).Left - 60)
    Next i
    
    '���ڴʾ������ĸ߶�
    For i = 1 To rtxtWord1.Count - 1
        ResizeRichTextBox rtxtWord1(i)
        If i = 1 Then
            rtxtWord1(i).Top = 30
        Else
            rtxtWord1(i).Top = rtxtWord1(i - 1).Top + rtxtWord1(i - 1).Height + 5
            
        End If
        chkSelect1(i).Top = rtxtWord1(i).Top
    Next
    
    lngH = 0
    If rtxtWord1.Count > 1 Then
        lngH = rtxtWord1(rtxtWord1.Count - 1).Top + rtxtWord1(rtxtWord1.Count - 1).Height + 200
    End If
    
    Call ResizeWordContainer(picCheckView, vscroWordH1, picContainer1, lngH)
End Sub


Private Sub picResult_Resize()
    Dim i As Integer
    Dim lngH As Long
    
    '����ÿһ��RichTextBox�Ŀ��
    For i = 1 To rtxtWord2.Count - 1
        rtxtWord2(i).Width = Abs(picContainer2.Width - rtxtWord2(i).Left - 60)
    Next i
    
    '���ڴʾ������ĸ߶�
    For i = 1 To rtxtWord2.Count - 1
        ResizeRichTextBox rtxtWord2(i)
        If i = 1 Then
            rtxtWord2(i).Top = 30
        Else
            rtxtWord2(i).Top = rtxtWord2(i - 1).Top + rtxtWord2(i - 1).Height + 5
        End If
        chkSelect2(i).Top = rtxtWord2(i).Top
    Next
    
    lngH = 0
    If rtxtWord2.Count > 1 Then
        lngH = rtxtWord2(rtxtWord2.Count - 1).Top + rtxtWord2(rtxtWord2.Count - 1).Height + 200
    End If
    
    Call ResizeWordContainer(picResult, vscroWordH2, picContainer2, lngH)
End Sub

Private Sub rtxtWord1_DblClick(Index As Integer)
    Call richTextBoxShowElements(rtxtWord1(Index))
End Sub

Private Sub rtxtWord2_DblClick(Index As Integer)
    Call richTextBoxShowElements(rtxtWord2(Index))
End Sub

Private Sub rtxtWord3_DblClick(Index As Integer)
    Call richTextBoxShowElements(rtxtWord3(Index))
End Sub

Private Sub vscroWordH1_Change()
    picContainer1.Top = -vscroWordH1.Value * 1000
End Sub

Private Sub vscroWordH2_Change()
    picContainer2.Top = -vscroWordH2.Value * 1000
End Sub

Private Sub vscroWordH3_Change()
    picContainer3.Top = -vscroWordH3.Value * 1000
End Sub
