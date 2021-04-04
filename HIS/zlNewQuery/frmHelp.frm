VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   5535
   ClientLeft      =   1845
   ClientTop       =   990
   ClientWidth     =   7875
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "�رհ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6465
      MouseIcon       =   "frmHelp.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   4920
      Width           =   1200
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00FFFFFF&
      Height          =   4800
      Left            =   75
      ScaleHeight     =   4740
      ScaleWidth      =   7665
      TabIndex        =   0
      Top             =   45
      Width           =   7725
      Begin VB.VScrollBar vsb 
         Height          =   3990
         Left            =   7020
         MouseIcon       =   "frmHelp.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   465
         Width           =   345
      End
      Begin VB.HScrollBar hsb 
         Height          =   330
         Left            =   90
         MouseIcon       =   "frmHelp.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   4350
         Width           =   2010
      End
      Begin VB.PictureBox picBack1 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   6360
         ScaleHeight     =   885
         ScaleWidth      =   1050
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   3780
         Width           =   1050
      End
      Begin zl9NewQuery.ctlQueryItem QueryItem 
         Height          =   2820
         Left            =   2100
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   375
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   4974
      End
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "���ǰ����鿴�������밴�ұߵ�[�رհ���]���˳���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   150
      TabIndex        =   6
      Top             =   5025
      Width           =   5520
   End
   Begin VB.Shape shp 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   5475
      Left            =   30
      Top             =   15
      Width           =   7800
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnFist As Boolean
Private mvarPageNo As Long
Private mvarSvrDept As String           '��������ҽ���Ŀ���
Private mvarSvrDuty As String           '��������ҽ����ְ��

Private mvarLeftStart As Single

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call InitInternal
End Sub

Private Sub Form_Activate()
    If mblnFist = False Then Exit Sub
    mblnFist = False
            
    DoEvents
    
    Call LoadPageItemList(mvarPageNo)
        
    Call CalcVsb
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call picBack_KeyDown(KeyCode, Shift)
End Sub

Private Sub Form_Load()
    mblnFist = True

    
    QueryItem.Height = Screen.Height
End Sub

Private Sub Form_Resize()
    '���ݴ���״̬,���������и��ؼ�����ʾλ��
    
    On Error Resume Next
    
    QueryItem.Width = Screen.Width - 2010 - 45
    Call ResizeControl(shp, 15, 15, Me.ScaleWidth - 30, Me.ScaleHeight - 30)
    
    Call ResizeControl(picBack, 45, 45, Me.ScaleWidth - 90, Me.ScaleHeight - cmdClose.Height - 120)
    Call ResizeControl(QueryItem, 0, 0, QueryItem.Width, QueryItem.Height)
    
    mvarLeftStart = QueryItem.Left
    
    Call ResizeControl(vsb, picBack.ScaleWidth - vsb.Width + 60, 0, vsb.Width, picBack.ScaleHeight - hsb.Height + 60)
    Call ResizeControl(hsb, 0, picBack.ScaleHeight - hsb.Height + 60, picBack.ScaleWidth - vsb.Width + 60, hsb.Height)
    picBack1.Left = vsb.Left
    picBack1.Top = hsb.Top
    
    Call ResizeControl(cmdClose, Me.ScaleWidth - cmdClose.Width - 60, picBack.Top + picBack.Height + 30, cmdClose.Width, cmdClose.Height)
    lbl.Top = cmdClose.Top + 75
    Call CalcVsb
End Sub

Public Function ShowHelp(frmMain As Object, ByVal PageNo As Long, ByVal vWidth As Single, ByVal vHeight As Single)
    mvarPageNo = PageNo
    frmHelp.Width = vWidth
    frmHelp.Height = vHeight
    frmHelp.Show 1, frmMain
End Function

Private Sub LoadPageItemList(ByVal PageNo As Long)
'����:����ҳ���ÿһ��ѯ��Ŀ
'����:PageNo            ҳ�����
'˵��:���ǲ�ѯ������ʾ�����岿��,��ʾ��ѯ����
    Dim FileName As String
    Dim W As Single
    Dim H As Single
    Dim vFont As New StdFont
    Dim i As Long
    Dim j As Long
    Dim strTmp As String
    Dim rs As New ADODB.Recordset
    Dim vNextY As Single
    Dim vNextX As Single
    Dim objDraw As ctlQueryItem
    Dim vWidth As Single
    Dim vHeight As Single
    Dim vTmp As Single
    Dim vTmp1 As Single
    Dim vMaxWidth As Single
    Dim vVisible As Boolean
    Dim strText As String
    
    On Error GoTo errHand
    i = 1
    vNextY = 60 + (i - 1) * 600
    vNextX = 120
    vMaxWidth = 120
            
    ShowFlatFlash "���Ժ���������ҳ��...", Me
    DoEvents
    
    Set objDraw = QueryItem
    objDraw.ClientVisible = False
    Call objDraw.ClearAllPageItem
    
    '��ȡҳ��ı������������
'    Set gRs = OpenRecord(gRs, "select B.����,B.���� from ��ѯҳ��Ŀ¼ A,��ѯͼƬԪ�� B where A.��������=B.��� and A.ҳ�����=" & PageNo)
'    If gRs.BOF = False Then FrameDefault.AdviceMovie = IIf(IsNull(gRs!����), "", App.Path & "\ͼ��\" & gRs!���� & IIf(gRs!���� <> 2, ".pic", ".swf"))
                    
    '��ʼ�����Զ����ѯҳ��
    gstrSQL = "select ҳ�����,�������,�����ı�,����ͼ��,��������,����λ��,��������,����ҳ��,��������,��������,������,���λ��,��ͼ���,��ͼλ�� from ��ѯ����Ŀ¼ where ҳ�����=[1] order by �������"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo)
    If gRs.BOF = False Then
        While Not gRs.EOF
            strTmp = IIf(IsNull(gRs!��������), "����;12;0;0;0", gRs!��������)
            vFont.Name = Split(strTmp, ";")(0)
            vFont.Size = Val(Split(strTmp, ";")(1))
            vFont.Bold = Val(Split(strTmp, ";")(2))
            vFont.Italic = Val(Split(strTmp, ";")(3))
                                    
            FileName = ""
            '1.���ر������ݼ�����ͼ��
            vVisible = IIf(IsNull(gRs!��������), 1, gRs!��������)
            
            gstrSQL = "select ���� from ��ѯͼƬԪ�� where ���=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(IIf(IsNull(gRs!����ͼ��), 0, gRs!����ͼ��)))
            If rs.BOF = False Then
                FileName = GetFileName(IIf(IsNull(gRs!����ͼ��), 0, gRs!����ͼ��), W, H)
            End If
            Call objDraw.AddPageItemTitle(i, vNextY, IIf(IsNull(gRs!�����ı�), "", gRs!�����ı�), Val(Split(strTmp, ";")(4)), vFont, FileName, PageNo, IIf(IsNull(gRs!�������), 0, gRs!�������), vWidth, vHeight, Not vVisible, IIf(IsNull(gRs!����λ��), 0, gRs!����λ��))
                                                                                    
            If Not vVisible = True Then vNextY = vNextY + vHeight + 150
            
            Select Case zlCommFun.Nvl(gRs("��������").Value, 0)
            '----------------------------------------------------------------------------------------------------------
            Case 0          '���ı�����
                strTmp = IIf(IsNull(gRs!��������), "����;12;0;0;0", gRs!��������)
                j = objDraw.NextTxtIndex
                
                vWidth = QueryItem.Width - 330
                
                strText = Sys.ReadLob(glngSys, 29, PageNo & "," & gRs("�������").Value, "", 1)
                Call objDraw.AddPageItemTxt(j, vNextX, vNextY, strText & Chr(13) & Chr(10), strTmp, vWidth, vHeight)
                vNextY = vNextY + vHeight + 150
            '----------------------------------------------------------------------------------------------------------
            Case 1          '���������
                vHeight = 0
                Call InsertGrid(objDraw, IIf(IsNull(gRs!������), 0, gRs!������), vNextX, vNextY, vWidth, vHeight)
                If vHeight > 0 Then vNextY = vNextY + vHeight + 150
            '----------------------------------------------------------------------------------------------------------
            Case 2          '��ͼ������
                FileName = GetFileName(IIf(IsNull(gRs!��ͼ���), 0, gRs!��ͼ���), W, H)
                Call objDraw.AddPageItemPic(objDraw.NextPicIndex, 0, vNextY, FileName, vWidth, vHeight, W, H)
                vNextY = vNextY + vHeight + 150
            '----------------------------------------------------------------------------------------------------------
            Case 3          '����������
                gstrSQL = "select C.ҳ������||decode(B.�����ı�,NULL,'','��'||B.�����ı�) as �����ı�,A.����ҳ��,A.ҳ�ڶκ� from ��ѯ�������� A,��ѯ����Ŀ¼ B,��ѯҳ��Ŀ¼ C Where A.����ҳ��=C.ҳ����� and A.����ҳ��=B.ҳ�����(+) and A.ҳ�ڶκ�=B.�������(+) and A.ҳ����� = [1] And A.������� = [2]"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo, Val(IIf(IsNull(gRs!�������), 0, gRs!�������)))
                If rs.BOF = False Then
                    While Not rs.EOF
                        Call objDraw.AddPageItemConnect(objDraw.NextConnectIndex, vNextX + 150, vNextY, IIf(IsNull(rs!�����ı�), "", rs!�����ı�), IIf(IsNull(rs!����ҳ��), 0, rs!����ҳ��), IIf(IsNull(rs!ҳ�ڶκ�), 0, rs!ҳ�ڶκ�), vWidth, vHeight)
                        vNextY = vNextY + 300
                        rs.MoveNext
                    Wend
                    vNextY = vNextY + 150
                Else
                    '����Ƿ����ӵ�ZLHIS����Ա
                    gstrSQL = "select B.����,A.����ҳ��,A.ҳ�ڶκ� from ��ѯ�������� A,��Ա�� B Where A.ҳ�ڶκ�=B.id And (b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.����ʱ�� Is Null) and A.ҳ����� = [1] And A.������� = [2]"
                    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo, Val(IIf(IsNull(gRs!�������), 0, gRs!�������)))
                    If rs.BOF = False Then
                        While Not rs.EOF
                            Call objDraw.AddPageItemConnect(objDraw.NextConnectIndex, vNextX, vNextY, IIf(IsNull(rs!����), "", rs!����), IIf(IsNull(rs!����ҳ��), 0, rs!����ҳ��), IIf(IsNull(rs!ҳ�ڶκ�), 0, rs!ҳ�ڶκ�), vWidth, vHeight)
                            vNextY = vNextY + 300
                            rs.MoveNext
                        Wend
                        vNextY = vNextY + 150
                    End If
                End If
            '----------------------------------------------------------------------------------------------------------
            Case 4          '�ı��ͱ��
                strTmp = IIf(IsNull(gRs!��������), "����;12;0;0;0", gRs!��������)
                j = objDraw.NextTxtIndex
                
                strText = Sys.ReadLob(glngSys, 29, PageNo & "," & gRs("�������").Value, "", 1)
                
                Select Case IIf(IsNull(gRs!���λ��), 0, gRs!���λ��)
                Case 0
                    vHeight = 0
                    Call InsertGrid(objDraw, IIf(IsNull(gRs!������), 0, gRs!������), 0, vNextY, vTmp1, vTmp)
                    vWidth = QueryItem.Width - vTmp1 - 120 - 120
                    Call objDraw.AddPageItemTxt(j, vNextX + vTmp1 + 60, vNextY, strText & Chr(13) & Chr(10), strTmp, vWidth, vHeight)
                Case 1
                    Call objDraw.AddPageItemTxt(j, vNextX, vNextY, strText & Chr(13) & Chr(10), strTmp, vTmp1, vHeight)
                    Call InsertGrid(objDraw, IIf(IsNull(gRs!������), 0, gRs!������), 1, vNextY, vWidth, vTmp)
                End Select
                vNextY = vNextY + IIf(vTmp > vHeight, vTmp, vHeight) + 150
            '----------------------------------------------------------------------------------------------------------
            Case 5          '�ı���ͼ��
                FileName = GetFileName(IIf(IsNull(gRs!��ͼ���), 0, gRs!��ͼ���), W, H)
                strText = Sys.ReadLob(glngSys, 29, PageNo & "," & gRs("�������").Value, "", 1)
                strTmp = IIf(IsNull(gRs!��������), "����;12;0;0;0", gRs!��������)
                j = objDraw.NextTxtIndex
                Select Case IIf(IsNull(gRs!��ͼλ��), 0, gRs!��ͼλ��)
                Case 0
                    Call objDraw.AddPageItemPic(objDraw.NextPicIndex, 0, vNextY, FileName, vTmp1, vTmp, W, H)
                    vWidth = QueryItem.Width - vTmp1 - 120 - 120
                    Call objDraw.AddPageItemTxt(j, vNextX + vTmp1 + 60, vNextY, strText & Chr(13) & Chr(10), strTmp, vWidth, vHeight)
                Case 1
                    Call objDraw.AddPageItemPic(objDraw.NextPicIndex, 1, vNextY, FileName, vWidth, vTmp, W, H)
                    vTmp1 = QueryItem.Width - vWidth - 60 - 90
                    Call objDraw.AddPageItemTxt(j, vNextX, vNextY, strText & Chr(13) & Chr(10), strTmp, vTmp1, vHeight)
                End Select
                vNextY = vNextY + IIf(vTmp > vHeight, vTmp, vHeight) + 150
            End Select
                        
            '8.���÷���ҳ�ױ�־
            If IIf(IsNull(gRs!����ҳ��), 0, gRs!����ҳ��) = 1 Then
                vHeight = 0
                Call objDraw.AddReturnFlag(vNextX, vNextY, vHeight)
                If vHeight > 0 Then vNextY = vNextY + vHeight + 150
            End If
            
            i = i + 1
            gRs.MoveNext
        Wend
    End If
        
    Call objDraw.ResizePage(QueryItem.Width, vNextY)
    QueryItem.Height = QueryItem.FactHeight
    'Call FrameDefault.InitNavigator(FrameDefault.ClientWidth, vNextY)
    
    '��ȡ����������ҳ�汳��
    gstrSQL = "select B.����,B.����,B.���,B.�߶� from ��ѯҳ��Ŀ¼ A,��ѯͼƬԪ�� B where A.ҳ�汳��=B.��� and A.ҳ�����=[1]"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo)
    If gRs.BOF = False Then
        Call objDraw.BackPicture(IIf(IsNull(gRs!����), "", App.Path & "\ͼ��\" & gRs!���� & IIf(gRs!���� <> 2, ".pic", ".swf")), IIf(IsNull(gRs!���), 0, gRs!���) * Screen.TwipsPerPixelX, IIf(IsNull(gRs!�߶�), 0, gRs!�߶�) * Screen.TwipsPerPixelY)
    End If
            
    Call objDraw.InitLoad
    objDraw.ClientVisible = True
    
    StopFlatFlash
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub hsb_Change()
    QueryItem.Left = mvarLeftStart - hsb.Value * 600
    If QueryItem.Left + QueryItem.Width < picBack.Left + picBack.Width - vsb.Width Then
        QueryItem.Left = picBack.Left + picBack.Width - QueryItem.Width - vsb.Width
    End If
    If QueryItem.Left > 0 Then QueryItem.Left = 0
    
End Sub

Private Sub hsb_KeyDown(KeyCode As Integer, Shift As Integer)
    Call picBack_KeyDown(KeyCode, Shift)
End Sub

Private Sub picBack_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If vsb.Enabled Then vsb.Value = IIf(vsb.Value < vsb.Max, vsb.Value + 1, vsb.Max)
    End If

    If KeyCode = vbKeyUp Then
        If vsb.Enabled Then vsb.Value = IIf(vsb.Value > 0, vsb.Value - 1, 0)
    End If

    If KeyCode = vbKeyRight Then
        If hsb.Enabled Then hsb.Value = IIf(hsb.Value < hsb.Max, hsb.Value + 1, hsb.Max)
    End If

    If KeyCode = vbKeyLeft Then
        If hsb.Enabled Then hsb.Value = IIf(hsb.Value > 0, hsb.Value - 1, 0)
    End If
End Sub

Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call InitInternal
End Sub

Private Sub picBack_Paint()
    Call RaisEffect(picBack, -1)
End Sub

Private Sub picBack1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call picBack_KeyDown(KeyCode, Shift)
End Sub

Private Sub picBack1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call InitInternal
End Sub

Private Sub QueryItem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call InitInternal
End Sub

Private Sub vsb_Change()
    QueryItem.Top = 0 - vsb.Value * 600
    If QueryItem.Top + QueryItem.Height < picBack.Top + picBack.Height - hsb.Height Then
        QueryItem.Top = picBack.Top + picBack.Height - hsb.Height - QueryItem.Height
    End If
    If QueryItem.Top > 0 Then QueryItem.Top = 0
    
End Sub

Private Sub CalcVsb()
    vsb.Max = 0 - Int(0 - (QueryItem.Height - picBack.ScaleHeight + hsb.Height + 45) / 600)
    If vsb.Max > 0 Then
        vsb.Enabled = True
        vsb.SmallChange = 1
        vsb.LargeChange = 1
        vsb.Value = 0
    Else
        vsb.Enabled = False
    End If
    
    hsb.Max = 0 - Int(0 - (QueryItem.Width - picBack.ScaleWidth + vsb.Width + 45) / 600)
    If hsb.Max > 0 Then
        hsb.Enabled = True
        hsb.SmallChange = 1
        hsb.LargeChange = 1
        hsb.Value = 0
    Else
        hsb.Enabled = False
    End If
End Sub

Private Sub vsb_KeyDown(KeyCode As Integer, Shift As Integer)
    Call picBack_KeyDown(KeyCode, Shift)
End Sub
