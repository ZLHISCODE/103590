VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPagePreview 
   Caption         =   "ҳ��Ԥ��"
   ClientHeight    =   6165
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   9330
   Icon            =   "frmPagePreview.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picBack 
      BackColor       =   &H8000000C&
      Height          =   4275
      Left            =   75
      ScaleHeight     =   4215
      ScaleWidth      =   5955
      TabIndex        =   0
      Top             =   45
      Width           =   6015
      Begin VB.PictureBox picBack1 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   4125
         ScaleHeight     =   495
         ScaleWidth      =   570
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   3240
         Width           =   570
      End
      Begin VB.HScrollBar hsb 
         Height          =   255
         Left            =   285
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   3870
         Width           =   2010
      End
      Begin VB.VScrollBar vsb 
         Height          =   3990
         Left            =   5670
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   -15
         Width           =   255
      End
      Begin zl9NewQuery.ctlQueryItem QueryItem 
         Height          =   2820
         Left            =   1365
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   285
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   4974
      End
   End
   Begin MSComctlLib.ImageList ilsMenuHot 
      Left            =   6960
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":06EA
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":090A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":0B2A
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":0D4A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":0F6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":14C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":1A1E
            Key             =   "View"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":1C3A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":1E5A
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   7545
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":207A
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":229A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":24BA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":26DA
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":28FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":2E54
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":33AE
            Key             =   "View"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":35CA
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPagePreview.frx":37EA
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPagePreview"
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

Private Sub Form_Activate()
    If mblnFist = False Then Exit Sub
    mblnFist = False
            
    DoEvents
    Call LoadPageItemList(mvarPageNo)
        
    Call CalcVsb
End Sub

Private Sub Form_Load()
    mblnFist = True
    RestoreWinState Me, App.ProductName
    
    QueryItem.Height = Screen.Height
End Sub

Private Sub Form_Resize()
    '���ݴ���״̬,���������и��ؼ�����ʾλ��

    QueryItem.Width = Screen.Width - 2010 - 45
    Call ResizeControl(picBack, 0, 0, Me.ScaleWidth, Me.ScaleHeight)
    Call ResizeControl(QueryItem, (Me.ScaleWidth - QueryItem.Width) / 2, 45, QueryItem.Width, QueryItem.Height)
    
    If QueryItem.Left < 45 Then QueryItem.Left = 45
    mvarLeftStart = QueryItem.Left
    
    Call ResizeControl(vsb, picBack.ScaleWidth - vsb.Width, 0, vsb.Width, picBack.ScaleHeight - hsb.Height)
    Call ResizeControl(hsb, 0, picBack.ScaleHeight - hsb.Height, picBack.ScaleWidth - vsb.Width, hsb.Height)
    picBack1.Left = vsb.Left
    picBack1.Top = hsb.Top
    
    Call CalcVsb
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Public Function ShowPreview(frmMain As Object, ByVal PageNo As Long)
    mvarPageNo = PageNo
    frmPagePreview.Show 1, frmMain
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
                gstrSQL = "select C.ҳ������||decode(B.�����ı�,NULL,'','��'||B.�����ı�) as �����ı�,A.����ҳ��,A.ҳ�ڶκ� from ��ѯ�������� A,��ѯ����Ŀ¼ B,��ѯҳ��Ŀ¼ C Where A.����ҳ��=C.ҳ����� and A.����ҳ��=B.ҳ�����(+) and A.ҳ�ڶκ�=B.�������(+) and A.ҳ����� = [1] And A.������� =[2] "
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
                    gstrSQL = "select B.����,A.����ҳ��,A.ҳ�ڶκ� from ��ѯ�������� A,��Ա�� B Where A.ҳ�ڶκ�=B.id And (b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.����ʱ�� Is Null) and A.ҳ����� = [1] And A.������� = [2] "
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
    On Error Resume Next
    QueryItem.Left = mvarLeftStart - hsb.Value * 600
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

Private Sub vsb_Change()
    On Error Resume Next
    QueryItem.Top = 45 - vsb.Value * 600
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
