VERSION 5.00
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelectRepImage 
   Caption         =   "��ȡ����ͼ�������ı�߿���ɫ�����ɫ�߿��ͼ��"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10290
   Icon            =   "frmSelectRepImage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   10290
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      Height          =   350
      Left            =   5640
      TabIndex        =   3
      Top             =   6120
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "��ȡ"
      Height          =   350
      Left            =   2640
      TabIndex        =   2
      Top             =   6120
      Width           =   1100
   End
   Begin DicomObjects.DicomViewer DViewer 
      Height          =   5535
      Left            =   3960
      TabIndex        =   1
      ToolTipText     =   "˫��������ʾ��ͼ"
      Top             =   0
      Width           =   6255
      _Version        =   262147
      _ExtentX        =   11033
      _ExtentY        =   9763
      _StockProps     =   35
      BackColor       =   0
   End
   Begin MSComctlLib.ListView lvwList 
      Height          =   5490
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   9684
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "dfd"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "dsd"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmSelectRepImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngOrderID As Long             '��ȡͼ���Ŀ��ҽ��ID
Private mlngSourceOrderID As Long       'Դҽ��id
Private mMultiRows As Integer
Private mMultiCols As Integer
Private mintSelectIndex As Integer
Private mImages As New DicomImages

'Private mlngShowBigImg As Long          '�Ƿ���ʾ��ͼ,0-����ʾ��1-����ƶ�ʱ��ʾ��2-��굥����ʾ��������
'Private mdblBigImgZoom As Double        '�����ͼ�Ŵ���

Private mExitState As Integer           '�˳�״̬   0-��ʾͨ��ȡ����ť�˳�   1-��ʾͨ��ȷ�����ͼ����˳�



Public Function ShowMe(frmParent As Form, lngOrderID As Long) As DicomImages
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim objItem As ListItem
    Dim iCount As Integer
    Dim strTime As String
    
    mlngOrderID = lngOrderID
''    mlngShowBigImg = lngShowBigImg
''    mdblBigImgZoom = dblBigImgZoom
    
    strSql = "Select c.Id As ҽ��id, Ӱ�����, c.����ʱ�� As ����ʱ��, c.ҽ������, b.����id " & _
            " From Ӱ�����¼ a, ����ҽ������ b, ����ҽ����¼ c " & _
            " Where a.ҽ��id = c.Id And b.ҽ��id = c.Id And c.����id = (Select ����id From ����ҽ����¼ Where Id = [1]) And " & _
            " c.���id Is Null And c.ִ�п���id =(Select ִ�п���id From ����ҽ����¼ Where Id = [1]) And c.Id <> [1] Order By ����ʱ�� Asc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ����ͼ��", mlngOrderID)
    
    lvwList.ListItems.Clear
    
    iCount = 1
    zlControl.LvwSelectColumns lvwList, "ҽ��ID,0,0,1;���,300,0,1;���,600,0,1;����ʱ��,1100,0,1;ҽ������,3000,0,1;����ID,0,0,1", True
    With lvwList
        Do While Not rsTemp.EOF
            Set objItem = .ListItems.Add(, "K" & rsTemp!ҽ��ID, rsTemp!ҽ��ID)
            '�������Ŀ
            objItem.SubItems(1) = iCount
            iCount = iCount + 1
            objItem.SubItems(2) = Nvl(rsTemp!Ӱ�����)
            strTime = Format(rsTemp!����ʱ��, "yyyy-mm-dd")
            objItem.SubItems(3) = strTime
            objItem.SubItems(4) = Nvl(rsTemp!ҽ������)
            objItem.SubItems(5) = Nvl(rsTemp!����Id)
            rsTemp.MoveNext
        Loop
    End With
    If lvwList.ListItems.Count > 0 Then
        Call lvwList_ItemClick(lvwList.SelectedItem)
    End If
    Me.Show 1, frmParent
    
    Set ShowMe = mImages
End Function

Private Sub cmdCancel_Click()
    mExitState = 0
    Unload Me
End Sub

Private Sub CmdOK_Click()
    Dim i As Integer
    Dim strFileNames As String
    Dim blnResult As Boolean
    Dim strSql As String
    
    'ת�Ʊ���ͼ��
    mImages.Clear
    If DViewer.Images.Count > 0 Then
        For i = 1 To DViewer.Images.Count
            If DViewer.Images(i).BorderColour = vbRed Then
                mImages.Add DViewer.Images(i)
            End If
        Next i
    End If
    
    mExitState = 1
    
    'ж�ش���
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mExitState = 0 Then mImages.Clear
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub DViewer_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)

    If DViewer.Images.Count = 0 Then Exit Sub
    
    If Button = 1 And Shift = 0 Then
        mintSelectIndex = DViewer.ImageIndex(X, Y)
        If DViewer.Images(mintSelectIndex).BorderColour = vbWhite Then
            DViewer.Images(mintSelectIndex).BorderColour = vbRed
        Else
            DViewer.Images(mintSelectIndex).BorderColour = vbWhite
        End If
    End If
End Sub

Private Sub DViewer_DblClick()
    
    If DViewer.Images.Count = 0 Then Exit Sub
    
    If DViewer.MultiColumns = 1 And DViewer.MultiRows = 1 Then
        DViewer.MultiColumns = mMultiCols
        DViewer.MultiRows = mMultiRows
        DViewer.CurrentIndex = 1
    Else
        mMultiCols = DViewer.MultiColumns
        mMultiRows = DViewer.MultiRows
        DViewer.MultiColumns = 1
        DViewer.MultiRows = 1
        DViewer.CurrentIndex = mintSelectIndex
    End If
End Sub

Private Sub DViewer_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
'    Dim blnShowImg As Boolean
'    Dim intCurrImg As Integer
'
'    If mlngShowBigImg = 0 Or DViewer.Images.Count <= 0 Then Exit Sub
'
'    '�ж��Ƿ���Ҫ��ʾͼ��
'    If (0 <= X * Screen.TwipsPerPixelX) And (X * Screen.TwipsPerPixelX <= DViewer.Width) And _
'       (0 <= Y * Screen.TwipsPerPixelY) And (Y * Screen.TwipsPerPixelY <= DViewer.Height) Then
'        blnShowImg = True
'    End If
'    If blnShowImg Then      '��ʾͼ��
'       SetCapture DViewer.hWnd    '�������
'
'        intCurrImg = DViewer.ImageIndex(X, Y)
'        If intCurrImg <> 0 Then
'            '����ͼ����ʾ
'            frmShowImg.ShowMe DViewer.Images(intCurrImg), Me, 1, 0, 0, mdblBigImgZoom
'        Else
'            frmShowImg.HideMe
'        End If
'    Else        '�ر�ͼ����ʾ
'        ReleaseCapture      '�������
'        frmShowImg.HideMe
'    End If
End Sub


Private Sub Form_Load()
    mExitState = 0
    Call RestoreWinState(Me, App.ProductName)
End Sub


Private Sub Form_Resize()
    lvwList.Left = 0
    lvwList.Top = 0
    lvwList.Width = Me.ScaleWidth * 0.4
    lvwList.Height = Abs(Me.ScaleHeight - 800)
    DViewer.Left = lvwList.Left + lvwList.Width + 50
    DViewer.Top = 0
    DViewer.Width = Me.ScaleWidth * 0.6
    DViewer.Height = lvwList.Height
    
    cmdOK.Left = Me.ScaleWidth * 0.3
    cmdOK.Top = Me.ScaleHeight - 600
    
    cmdCancel.Left = Me.ScaleWidth * 0.7
    cmdCancel.Top = cmdOK.Top
End Sub



Private Sub lvwList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    zlControl.LvwSortColumn lvwList, ColumnHeader.Index
End Sub

Private Sub lvwList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer
    
    mlngSourceOrderID = Mid(Item.Key, 2)
    
    '��ʾ����ͼ
    Call GetRptImages(Me.DViewer, mlngSourceOrderID, False)
    
    mMultiCols = 1
    mMultiRows = 1
    
    For i = 1 To DViewer.Images.Count
        DViewer.Images(i).BorderColour = vbWhite
    Next i
    
    If DViewer.Images.Count > 0 Then
        mintSelectIndex = 1
    Else
        mintSelectIndex = 0
    End If
End Sub
