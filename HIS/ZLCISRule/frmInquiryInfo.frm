VERSION 5.00
Begin VB.Form frmInquiryInfo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   9090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10830
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picBottom 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   120
      ScaleHeight     =   720
      ScaleWidth      =   9975
      TabIndex        =   7
      Top             =   8280
      Width           =   9975
      Begin VB.PictureBox picBtn 
         Appearance      =   0  'Flat
         BackColor       =   &H00D48A00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   0
         Left            =   4440
         ScaleHeight     =   420
         ScaleWidth      =   1095
         TabIndex        =   9
         Top             =   120
         Width           =   1100
         Begin VB.Label lblBtn 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ȷ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   0
            Left            =   360
            TabIndex        =   10
            Top             =   120
            Width           =   450
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   9960
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   240
      ScaleHeight     =   4815
      ScaleWidth      =   9975
      TabIndex        =   2
      Top             =   1080
      Width           =   9975
      Begin VB.Frame fraTable 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   3015
         Left            =   4680
         TabIndex        =   12
         Top             =   840
         Width           =   3855
      End
      Begin VB.VScrollBar vsc 
         Height          =   3360
         Left            =   9600
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox txtItem 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   3960
         Visible         =   0   'False
         Width           =   6135
      End
      Begin VB.ListBox lstItem 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   0
         Left            =   360
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   2280
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.ComboBox cboItem 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1200
         Visible         =   0   'False
         Width           =   3855
      End
   End
   Begin VB.Timer tmrTime 
      Interval        =   50
      Left            =   6600
      Top             =   7680
   End
   Begin VB.PictureBox picTop 
      BackColor       =   &H00D48A00&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   9975
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      Begin VB.PictureBox picClosed 
         Appearance      =   0  'Flat
         BackColor       =   &H00D48A00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   500
         Left            =   8520
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   3
         Top             =   0
         Width           =   500
         Begin VB.Label lblClose 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   300
            Left            =   120
            TabIndex        =   4
            Top             =   120
            Width           =   300
         End
      End
      Begin VB.Label lblFrmName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������Ϣ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   900
      End
   End
   Begin VB.Line linScope 
      Index           =   3
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   10560
   End
   Begin VB.Line linScope 
      Index           =   1
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   10560
   End
   Begin VB.Line linScope 
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   10560
   End
   Begin VB.Line linScope 
      Index           =   2
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   10560
   End
End
Attribute VB_Name = "frmInquiryInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mMoveX As Long, mMoveY As Long  '��¼�����ƶ�ǰ���������Ͻ������ָ��λ�ü���ݺ����
Private mudtRect As RECT
Private mudtRectClose As RECT

Private mudtPoint As POINTAPI
Private mblnMoveStart As Boolean '�ж��ƶ��Ƿ�ʼ
Private mblnMove As Boolean
   
Private mcolList As Collection
Private mstrJson As String

Public Function ShowMe(ByVal colList As Collection, ByRef strJsonOut As String) As Boolean
'����: colList ���ϵļ���,ÿ������������String���ͼ�һ���������͹��ɡ�
'     ����String����: observ_item_id|item_name|item_code
'     һ����������:  observ_item_values(����Ԫ��:item_detail_id|disp_name|default_sign)
    Set mcolList = colList
    mstrJson = ""
    Me.Show 1
    strJsonOut = mstrJson
    ShowMe = True
End Function

Private Sub Form_Activate()
    glngOldWindowProc = GetWindowLong(vsc.hWnd, GWL_WNDPROC)
    '��vsc����Ϣ������ָ��Ϊ�Զ��庯��NewWindowProc;ͬʱ��¼��ԭ��Ϣ��������ַ
    glngOldWindowProc = SetWindowLong(vsc.hWnd, GWL_WNDPROC, AddressOf NewWindowProc)
End Sub

Private Sub Form_Deactivate()
    '��WindowsĬ�ϵĺ����������¼�
    Call SetWindowLong(vsc.hWnd, GWL_WNDPROC, glngOldWindowProc)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCur As Long, lngMin As Long, lngMax As Long
    
    If vsc.Visible Then
        lngCur = vsc.Value
        lngMin = vsc.Min
        lngMax = vsc.Max
        If lngMax <= lngMin Then '��ֱ������δ����
            If KeyCode = vbKeyPageDown Then '��
                If Between(lngCur + (lngMax - lngMin) / 10, lngMin, lngMax) Then
                    vsc.Value = lngCur + (lngMax - lngMin) / 10
                Else
                    vsc.Value = lngMax
                End If
            ElseIf KeyCode = vbKeyPageUp Then '��
                If Between(lngCur - (lngMax - lngMin) / 10, lngMin, lngMax) Then
                    vsc.Value = lngCur - (lngMax - lngMin) / 10
                Else
                    vsc.Value = lngMin
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = Asc("`") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    
    picTop.BackColor = conCOLOR_TITLE_BAR
    vsc.Width = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX
    Me.Width = 12000: Me.Height = 9000
    If Not LoadForm Then Unload Me: Exit Sub
End Sub

Private Function LoadForm() As Boolean
      '����:�������ݶ�̬���ؿؼ�
      '���ݽ�����\�߶ȼ���ؼ�λ��
      'ListBox ��Щ������ֻ������,��Controls.Add ��ʽ�������ܶ��丳ֵ,�ʲ��ÿؼ����鷽ʽʵ��
      'ע�⣺��̬���صĿؼ������ڹؼ���,�޸�ʱע�ⲻҪ��©,��������ʱ����Name���ж�
          Dim i As Long, j As Long
          Dim colValues As Collection
          Dim lngH As Long, lngRowHeight As Long, lngSplitX As Long, lngSplitY As Long
          Dim lngTempX As Long, lngTempY As Long, lngTempW As Long
          Dim bytFontSize As Byte
          

          Dim strName As String
          
          Dim fraItem As VB.Frame
          Dim fraGroup As VB.Frame
          
          Dim lblItem As VB.Label
          Dim lblGroup As VB.Label
          Dim optItem As VB.OptionButton
          Dim lngGroupH As Long
          Dim colTemp  As Collection
          
          '��������
1         On Error GoTo ErrH

2         lngH = 0          '�ܸ߶�
3         lngSplitX = 120     'ˮƽ���
4         lngSplitY = 120     '��ֱ���
5         lngRowHeight = 300 '�����и�
6         bytFontSize = 11.5
7         Me.FontSize = bytFontSize
          
8         For i = 1 To mcolList.Count
              '��ӷ����
9             If Not fraGroup Is Nothing Then
10                fraGroup.Height = lngGroupH
11                lngH = lngH + lngGroupH + lngSplitY
12                lngGroupH = 0
13            End If
14            SetContolAttr "Frame", fraGroup, "fraGroup" & "_" & i, "", fraTable, , , , 1
15            fraGroup.Move 0, lngH, fraTable.Width
16            Call GetSubString(fraGroup.Width - lngSplitX * 2, mcolList(i)("item_name"), colTemp)
17            lngTempY = lngSplitX * 2
18            For j = 1 To colTemp.Count
19                SetContolAttr "Label", lblGroup, "lblGroup" & "_" & i & "_" & j, colTemp(j), fraGroup, bytFontSize
20                lblGroup.Move lngSplitX, lngTempY
21                lngTempY = lblGroup.Top + lblGroup.Height + lngSplitY
22            Next
23            lngGroupH = lngTempY + lngSplitY
                  
              'ÿ����¼����Ӧһ��Frame
24            Set colValues = mcolList(i)("observ_item_values")
25            SetContolAttr "Frame", fraItem, "fraItem" & "_" & i, "", fraGroup
26            fraItem.Height = lngRowHeight
27            fraItem.Move lngSplitX, lngGroupH, fraGroup.Width - lngSplitX * 2, lngRowHeight
           
28            lngTempX = 0: lngTempY = 0
29            For j = 1 To colValues.Count
30                Set optItem = Controls.Add("VB.OptionButton", "optBtn" & "_" & i & "_" & j, fraItem)
31                strName = GetCollValue(colValues, j, "disp_name")
32                optItem.Visible = True
33                optItem.FontSize = bytFontSize
34                optItem.Caption = strName
35                optItem.Tag = GetCollValue(colValues, j, "item_detail_id") 'ID
                   
                  'Ĭ��ֵ����
36                optItem.Value = (Val(GetCollValue(colValues, j, "default_sign")) = 1)
                   
37                optItem.Height = TextHeight(strName)
38                optItem.BackColor = fraItem.BackColor
                   
                  '��ѡ��Ŀ���
39                lngTempW = TextWidth("AAA") + TextWidth(strName)
                  If lngTempX + lngTempW > fraItem.Width Then
                      '�����߽绻��
                      lngTempY = lngTempY + optItem.Height + lngSplitY
                      lngTempX = 0
                  End If
40                optItem.Move lngTempX, lngTempY, lngTempW
41                lngTempX = lngTempX + lngTempW + lngSplitX '��¼��һ����ĿLEFTֵ
42            Next
43            strName = GetCollElement(mcolList, "item_name")
44            Call GetSubString(fraItem.Width - lngTempX, strName, colTemp)
45            For j = 1 To colTemp.Count
46                SetContolAttr "Label", lblItem, "lblItem" & "_" & i & "_" & j, colTemp(j), fraItem, bytFontSize
47                lblItem.Move lngTempX, lngTempY
48                lngTempY = lblItem.Top + lblItem.Height + lngSplitX
49            Next
50            fraItem.Height = lngTempY
          
              '��¼��һ��λ��
51            lngGroupH = fraItem.Top + fraItem.Height + lngSplitY
52            If mcolList.Count = i Then
53                fraGroup.Height = lngGroupH
54                lngH = lngH + lngGroupH + lngSplitY
55            End If
56        Next
           
57        fraTable.Height = lngH
58        If fraTable.Height < picMain.Height Then
59            vsc.Visible = False
60            vsc.Tag = "0"
61        Else
62            vsc.Tag = "1"
63            vsc.Visible = True
64            vsc.Value = 0
65            vsc.Min = 0
66            vsc.Max = (picMain.ScaleHeight - fraTable.Height) / Screen.TwipsPerPixelY
67            vsc.SmallChange = 5
68            vsc.LargeChange = 50
69            Me.Width = Me.ScaleWidth + vsc.Width
70        End If
71        LoadForm = True
72        Exit Function

ErrH:
73        MsgBox "��LoadForm�ĵ�" & Erl() & "�г���" & vbCrLf & _
            "�����: " & Err.Number & vbCrLf & _
            "����������" & Err.Description, vbExclamation, "frmInquiryInfo"
End Function

Private Sub GetSubString(ByVal lngLen As Long, ByVal strSource As String, ByRef colStr As Collection)
'����:��ָ�����Ƚ�ȡ�ַ�
    Dim lngMid As Long
    Dim lngMin As Long, lngMax As Long
    Set colStr = New Collection
    If TextWidth(strSource) < lngLen Then colStr.Add strSource: Exit Sub
    
    Do While strSource <> ""
        lngMin = 1: lngMax = Len(strSource)
        Do While lngMin <= lngMax
            lngMid = (lngMin + lngMax) \ 2
            If TextWidth(Mid(strSource, 1, lngMid)) > lngLen Then
                lngMax = lngMid - 1
            ElseIf TextWidth(Mid(strSource, 1, lngMid)) < lngLen Then
                lngMin = lngMid + 1
            Else
                Exit Do
            End If
        Loop
        colStr.Add Mid(strSource, 1, lngMid)
        strSource = Mid(strSource, lngMid + 1)
        If strSource = "" Then Exit Do
        If TextWidth(strSource) < lngLen Then
            colStr.Add strSource
            strSource = ""
        End If
    Loop
End Sub

Private Sub lblBtn_Click(index As Integer)
          Dim i As Long, j As Long
          Dim optItem As VB.OptionButton
          Dim strMsg As String
          Dim colValues As Collection
          Dim strJson As String
          Dim strItemId As String
          Dim strItemName As String
          
          '��֯����
1         On Error GoTo ErrH

2         For i = 1 To mcolList.Count
3             Set colValues = mcolList(i)("observ_item_values")
4             If Not colValues Is Nothing Then
5                 For j = 1 To colValues.Count
6                     Set optItem = Controls.Item("optBtn" & "_" & i & "_" & j)
7                     If optItem.Value = True Then
8                         strItemId = optItem.Tag
9                         strItemName = optItem.Caption
10                        Exit For
11                    ElseIf j = colValues.Count Then
12                        strMsg = strMsg & "��" & GetCollValue(mcolList, i, "item_name") & "��δ��ѡ��" & vbCrLf
13                    End If
14                Next
15            End If
16            If strJson <> "" Then strJson = strJson & ","
17            strJson = strJson & "{'observ_item_id':'" & GetCollValue(mcolList, i, "observ_item_id") & _
                                  "','item_name':'" & GetCollValue(mcolList, i, "item_name") & _
                                  "','item_code':'" & GetCollValue(mcolList, i, "item_code") & _
                                  "','item_detail_id':'" & strItemId & _
                                  "','disp_name':'" & strItemName & "'}"
18        Next
19        strJson = "'inquiry':[" & strJson & "]"
20        mstrJson = Replace(strJson, "'", "\""")
21        If strMsg <> "" Then
22            MsgBox strMsg, vbExclamation + vbOKOnly, gstrSysName
23            Exit Sub
24        End If
25        Unload Me

26        Exit Sub

ErrH:
27        MsgBox "��lblBtn_Click�ĵ�" & Erl() & "�г���" & vbCrLf & _
                  "�����: " & Err.Number & vbCrLf & _
                  "����������" & Err.Description, vbExclamation, "frmInquiryInfo"
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    picTop.Move 15, 15, Me.ScaleWidth - 30, 495
    picBottom.Move 15, Me.ScaleHeight - 720, picTop.Width, 705
    picMain.Move 15, picTop.Height + 15, Me.ScaleWidth - 30, Me.ScaleHeight - picTop.Height - picBottom.Height - 30
    If vsc.Tag = "1" Then
        vsc.Move picMain.ScaleWidth - vsc.Width, 0, vsc.Width, picMain.Height - 60
    End If
    fraTable.Move 120, 0, picMain.Width - IIf(vsc.Tag = "1", vsc.Width, 0) - 240
    
    picBottom.BackColor = picMain.BackColor
    'Left
    With linScope(0)
        .X1 = 0: .X2 = 0: .Y1 = 0: .Y2 = Me.ScaleHeight
        .BorderColor = conCOLOR_TITLE_BAR
    End With
    'bottom
    With linScope(1)
        .X1 = 0: .X2 = Me.ScaleWidth: .Y1 = Me.ScaleHeight - 15: .Y2 = Me.ScaleHeight - 15
        .BorderColor = conCOLOR_TITLE_BAR
    End With
    'right
    With linScope(2)
        .X1 = Me.ScaleWidth - 15: .X2 = Me.ScaleWidth - 15: .Y1 = 0: .Y2 = Me.ScaleHeight - 15
        .BorderColor = conCOLOR_TITLE_BAR
    End With
    'Top
    With linScope(3)
        .X1 = 0: .X2 = Me.ScaleWidth - 15: .Y1 = 0: .Y2 = 0
        .BorderColor = conCOLOR_TITLE_BAR
    End With
End Sub

Private Sub lblClose_Click()
    Call lblBtn_Click(0)
End Sub

Private Sub picBottom_Resize()
    On Error Resume Next
    picBtn(0).Move picBottom.Width / 2 - picBtn(0).Width / 2, picBottom.Height / 2 - picBtn(0).Height / 2
    With Line1
        .X1 = 120: .Y1 = 0
        .X2 = picBottom.ScaleWidth - 240: .Y2 = 0
        .BorderColor = conCOLOR_TITLE_BAR
    End With
End Sub

Private Sub picBtn_Click(index As Integer)
    Call lblBtn_Click(0)
End Sub

Private Sub picClosed_Click()
    Call lblBtn_Click(0)
End Sub

Private Sub picClosed_Resize()
    On Error Resume Next
    lblClose.Move picClosed.ScaleWidth / 2 + lblClose.Width / 2, (picClosed.ScaleHeight - lblClose.Height) / 2
End Sub

Private Sub picTop_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If mblnMove Then
        mMoveX = mudtPoint.x - mudtRect.Left
        mMoveY = mudtPoint.Y - mudtRect.Top
        mblnMoveStart = True
    End If
End Sub

Private Sub picTop_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim lngRet As Long
    If mblnMoveStart Then
        lngRet = MoveWindow(Me.hWnd, mudtPoint.x - mMoveX, mudtPoint.Y - mMoveY, mudtRect.Right - mudtRect.Left, mudtRect.Bottom - mudtRect.Top, -1)
    End If
End Sub

Private Sub picTop_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call GetWindowRect(Me.hWnd, mudtRect)
    Call GetWindowRect(picClosed.hWnd, mudtRectClose)
    mblnMoveStart = False
End Sub

Private Sub picTop_Resize()
    On Error Resume Next
    lblFrmName.Move 120, picTop.ScaleHeight / 2 - lblFrmName.Height / 2
    picClosed.Move picTop.ScaleWidth - picTop.Height, 0, picTop.Height, picTop.Height
End Sub
 
Private Sub tmrTime_Timer()
    Dim lngRet As Long
    If tmrTime.Tag = "" Then
        Call GetWindowRect(Me.hWnd, mudtRect)
        Call GetWindowRect(picClosed.hWnd, mudtRectClose)
        tmrTime.Tag = "1" '�״μ�¼����λ��
    End If
    lngRet = GetCursorPos(mudtPoint)
    '�ж����ָ���Ƿ�λ�ڴ����϶���
    If PtInRect(mudtRect, mudtPoint.x, mudtPoint.Y) Then
       mblnMove = True
    Else
       mblnMove = False
    End If
    If PtInRect(mudtRectClose, mudtPoint.x, mudtPoint.Y) Then
        picClosed.BackColor = "&H" & Hex(RGB(212, 64, 39))  '��ɫ
    Else
        picClosed.BackColor = picTop.BackColor
    End If
End Sub

Private Sub vsc_Change()
    vsc_Scroll
End Sub

Private Sub vsc_Scroll()
    fraTable.Top = vsc.Value * Screen.TwipsPerPixelY
    If vsc.Enabled Then vsc.SetFocus
End Sub

Private Sub SetContolAttr(ByVal strCtlType As String, ByRef objCtl As Object, ByVal strCtlName As String, ByVal strCaption As String, ByRef objContainer As Object, _
    Optional ByVal bytFontSize As Byte, Optional ByVal blnVisible As Boolean = True, _
    Optional ByVal blnAutoSize As Boolean = True, Optional ByVal bytBorderStyle As Byte)
    Select Case strCtlType
    
    Case "Label"
        Set objCtl = Controls.Add("VB.Label", strCtlName, objContainer)
        objCtl.FontSize = bytFontSize
        objCtl.Visible = blnVisible
        objCtl.AutoSize = blnAutoSize
        objCtl.Caption = strCaption
        objCtl.BackColor = objContainer.BackColor
    Case "Frame"
        Set objCtl = Controls.Add("VB.Frame", strCtlName, objContainer)
        objCtl.Visible = blnVisible
        objCtl.BorderStyle = bytBorderStyle
        objCtl.BackColor = objContainer.BackColor
        objCtl.Caption = strCaption
    Case "CheckBox"
        Set objCtl = Controls.Add("VB.CheckBox", strCtlName, objContainer)
        objCtl.Visible = blnVisible
        objCtl.FontSize = bytFontSize
        objCtl.Caption = strCaption
        objCtl.BackColor = objContainer.BackColor
    Case "ListBox"
        objCtl.FontSize = bytFontSize
        objCtl.Visible = blnVisible
        Set objCtl.Container = objContainer
    End Select
                
End Sub

'Private Sub SavePatiStatus(ByVal rsAsk As ADODB.Recordset)
''����:����״̬����
'    Dim strJson         As String
'    Dim strPvid         As String
'    Dim strenvr_id      As String
'    Dim strNo           As String
'    Dim strCurrTime     As String
'    Dim strVisitTime    As String
'    Dim strVisitDoc     As String
'    Dim strStatus       As String
'    Dim strOut          As String
'    Dim strErr          As String
'    Dim bytVType        As Byte
'    Dim rsPati          As ADODB.Recordset
'    Dim i               As Long
'
'    If gstrStatusSave = "" Then Exit Sub
'    If rsAsk Is Nothing Then Exit Sub
'    'http://192.168.0.231:8080/ords/patstatus/pat/saverecord
'    Set rsPati = GetPatiInfo_YF(gobjPati.lng����ID, gobjPati.str�Һŵ�, gobjPati.lng��ҳID)
'    If glngModel = PM_����༭ Then
'        strenvr_id = "10"
'    ElseIf glngModel = PM_סԺ�༭ Then
'        strenvr_id = "11"
'    End If
'    If gobjPati.str�Һŵ� <> "" Then
'        strPvid = gobjPati.str�Һŵ�
'        bytVType = 1
'        strNo = rsPati!����� & ""
'        strVisitTime = Format(rsPati!����ʱ��, "YYYY-MM-DD HH:MM:SS")
'        strVisitDoc = rsPati!ִ���� & ""
'    Else
'        strPvid = gobjPati.lng��ҳID & ""
'        bytVType = 2
'        strNo = rsPati!סԺ�� & ""
'        strVisitTime = Format(rsPati!��Ժʱ��, "YYYY-MM-DD HH:MM:SS")
'        strVisitDoc = ""
'    End If
'    strCurrTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
'
'    strJson = "{""rec_info"":[{""pid"":""" & gobjPati.lng����ID & """,""name"":""" & gobjPati.str���� & """," & _
'            """sex"":""" & gobjPati.str�Ա� & """,""birth"":""" & Format(rsPati!�������� & "", "YYYY-MM-DD") & """," & _
'            """age"":""" & rsPati!���� & """,""pvid"":""" & strPvid & """," & _
'            """visit_type"":""" & bytVType & " "",""envr_id"":""" & strenvr_id & """," & _
'            """visit_identifier"":""" & strNo & """,""visit_time"":""" & strVisitTime & """," & _
'            """marry_cnds"":""" & rsPati!����״�� & """,""visit_dept"":""" & rsPati!��ǰ���� & """," & _
'            """visit_doc"":""" & strVisitDoc & """,""rec_time"":""" & strCurrTime & """," & _
'            """recorder"":""" & UserInfo.���� & """,""recorder_id"":""" & UserInfo.id & """}]," & vbNewLine & _
'            """rec_detail"":["
'
'    '״̬ID �ɲ��� 'status_situation 1-������;3-�����
'    For i = 1 To rsAsk.RecordCount
'        strStatus = strStatus & ",{""strtus_id"":"""",""status_name"":""" & rsAsk!Index & """,""status_situation"":""" & IIf(rsAsk!Default = "��", 3, 1) & """}"
'        rsAsk.MoveNext
'    Next
'    If strStatus <> "" Then
'        strStatus = Mid(strStatus, 2)
'    Else
'        strStatus = "{""strtus_id"":"""",""status_name"":"""",""status_situation"":""""}"
'    End If
'    strJson = strJson & strStatus & "]}"
'    WriteLog "" & glngModel, "SavePatiStatus", "����״̬����URL:" & gstrStatusSave & ",����ֵ:" & strJson
'    Call sys.WebAPIByBasic(gstrStatusSave, strJson, strOut, strErr)
'    WriteLog "" & glngModel, "SavePatiStatus", "����״̬���� ����ֵ:" & strOut & IIf(strErr <> "", "������Ϣ:" & strErr, "")
'End Sub


