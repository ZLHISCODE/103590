VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm��ͨ��ѯ���� 
   Caption         =   "��ͨ��ѯ����"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11655
   Icon            =   "Frm��ѯ����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgColor 
      Left            =   10080
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm��ѯ����.frx":06EA
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm��ѯ����.frx":0904
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm��ѯ����.frx":0B1E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm��ѯ����.frx":0D38
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm��ѯ����.frx":0F52
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm��ѯ����.frx":164C
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm��ѯ����.frx":1866
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm��ѯ����.frx":1A80
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   10800
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm��ѯ����.frx":1C9A
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm��ѯ����.frx":1EB4
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm��ѯ����.frx":20CE
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm��ѯ����.frx":22E8
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm��ѯ����.frx":2502
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm��ѯ����.frx":2BFC
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm��ѯ����.frx":2E16
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm��ѯ����.frx":3030
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6390
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "Frm��ѯ����.frx":324A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15505
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh��ϸ_S 
      Height          =   5655
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   9975
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      MouseIcon       =   "Frm��ѯ����.frx":3ADC
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   1244
      BandCount       =   1
      ForeColor       =   -2147483635
      FixedOrder      =   -1  'True
      _CBWidth        =   11655
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      BandForeColor1  =   -2147483635
      Child1          =   "tbrThis"
      MinWidth1       =   4995
      MinHeight1      =   645
      Width1          =   810
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Begin VB.TextBox Txt��ˮ�� 
         Height          =   375
         Left            =   8880
         TabIndex        =   4
         Text            =   "�����س�"
         Top             =   120
         Width           =   2655
      End
      Begin VB.TextBox Txtҵ������ 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   375
         Left            =   7080
         TabIndex        =   3
         Text            =   "ҵ������"
         Top             =   120
         Width           =   1695
      End
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   1138
         ButtonWidth     =   820
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Description     =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFileSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileCard 
         Caption         =   "��Ƭ��ӡ(&A)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSplit2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuBusinessed 
         Caption         =   "��������"
      End
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewStatus 
            Caption         =   "״̬��(&S)"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&WEB�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frm��ͨ��ѯ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mtxtCaption As String, mBusinessedDay As Integer
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Sub cbr_HeightChanged(ByVal NewHeight As Single)
    Call ResizeForm
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If InStr("1234567890" & Chr(0) & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error Resume Next
    cbr.Height = 360
    mBusinessedDay = Val(GetSetting(appName:="ZLSOFT", Section:="˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, Key:="��������", Default:=30))
    Call InitTable
    RestoreWinState Me, App.ProductName
End Sub
Private Sub InitTable()
    Select Case mtxtCaption
    Case "������ˮ��"
        Txtҵ������.Text = "����id"
        With msh��ϸ_S
            .Clear
            .Rows = 2
            .Cols = 8
            .TextMatrix(0, 0) = "������ˮ��"
            .TextMatrix(0, 1) = "��ϸ��"
            .TextMatrix(0, 2) = "����"
            .TextMatrix(0, 3) = "����"
            .TextMatrix(0, 4) = "��λ"
            .TextMatrix(0, 5) = "����"
            .TextMatrix(0, 6) = "����"
            .TextMatrix(0, 7) = "���"
            .ColWidth(0) = 1800
            .ColWidth(1) = 600
            .ColWidth(2) = 900
            .ColWidth(3) = 2200
            .ColWidth(4) = 400
            .ColWidth(5) = 300
            .ColWidth(6) = 600
            .ColWidth(7) = 600
        End With
    Case "סԺ��ˮ��"
        Txtҵ������.Text = "����id"
        With msh��ϸ_S
            .Clear
            .Rows = 2
            .Cols = 12
        End With
    End Select
    Txt��ˮ��.SetFocus
End Sub
Private Sub Form_Resize()
    Call ResizeForm
End Sub

Private Sub ResizeForm()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    With msh��ϸ_S
        .Top = IIf(cbr.Visible, cbr.Height, 0)
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With

    tbrThis.Width = Me.ScaleWidth - Txtҵ������.Width - Txt��ˮ��.Width - Txtҵ������.Width \ 3
    Txtҵ������.Left = tbrThis.Width
    Txt��ˮ��.Left = Txtҵ������.Left + Txtҵ������.Width
End Sub
Private Sub Form_Unload(Cancel As Integer)
    mtxtCaption = ""
    mBusinessedDay = 0
    SaveWinState Me, App.ProductName
End Sub

Private Sub mnuBusinessed_Click()
Dim Businessed As String
    Businessed = Trim(InputBox("�������������", gstrSysName, mBusinessedDay))
    If Businessed <> "" Then
        mBusinessedDay = Businessed
    End If
    Call SaveSetting(appName:="ZLSOFT", Section:="˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, Key:="��������", setting:=mBusinessedDay)
End Sub

Private Sub mnuHelpTitle_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub
Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFileQuit_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnuFilePreview_Click
        Case "Help"
            mnuHelpTitle_Click
    End Select
End Sub

Private Sub mnuFileQuit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub
Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = Not stbThis.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    cbr.Visible = Not cbr.Visible
    Form_Resize
End Sub

Private Sub tbrThis_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub mnuFileExcel_Click()
    Call subPrint(3)
End Sub

Private Sub mnuFilePreview_Click()
    Call subPrint(2)
End Sub

Private Sub mnuFilePrint_Click()
    Call subPrint(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub subPrint(bytStyle As Byte)
'���ܣ�������б�
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    
    If gstrUserName = "" Then GetUserInfo
    intRow = msh��ϸ_S.Row
    
    '��ͷ
    objOut.Title.Text = Me.Caption
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    'objRow.Add "ҽ�����" & cmb����.Text
    'objOut.UnderAppRows.Add objRow
    
    'Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate, "yyyy��MM��DD��")
    objOut.BelowAppRows.Add objRow
    
    '����
    Set objOut.Body = msh��ϸ_S
    
    '���
    msh��ϸ_S.Redraw = False
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    msh��ϸ_S.Redraw = True
    
    msh��ϸ_S.Row = intRow
    msh��ϸ_S.Col = 0: msh��ϸ_S.ColSel = msh��ϸ_S.Cols - 1
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage Me.hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo Me.hwnd
End Sub
Public Sub ShowForm(frmCaption As String, txtCaption As String)
mtxtCaption = txtCaption
Me.Caption = frmCaption
Me.Show 1
End Sub

Private Sub Txt��ˮ��_GotFocus()
Txt��ˮ��.Text = ""
End Sub

Private Sub Txt��ˮ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Dim rsTemp As ADODB.Recordset, vRect As RECT, blncancle As Boolean
        If Trim(Txt��ˮ��) = "" Then Exit Sub
        vRect = GetControlRect(Txtҵ������.hwnd)
        Select Case mtxtCaption
        Case "������ˮ��"
            '�����ֶ�����¼ʱ���ֵ�����,����ȡ���㽻�׺�
            gstrSQL = "select C.����id AS ID ,A.֧��˳���,D.����,A.�������ý��,A.�����ʻ�֧��,to_char(C.�տ�ʱ��,'yyyy-mm-dd hh24:mi:ss') as �շ�ʱ�� from ���ս����¼ A ,������Ϣ D," & _
                            "(select distinct ����id,�տ�ʱ��,����id from ����Ԥ����¼ where ��¼����=3 and ��¼״̬=1 and ����id=" & Txt��ˮ��.Text & ") C " & _
                    "where C.����id=A.��¼id and a.����=1 and a.����=" & TYPE_��ͨ & " and d.����id=a.����id and C.�տ�ʱ��>=to_date('" & Format(zlDatabase.Currentdate - mBusinessedDay, "yyyy-mm-dd HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')"
            Set rsTemp = zlDatabase.ShowSelect(Me, gstrSQL, 0, , , , , , , True, vRect.Left - Txt��ˮ��.Width, vRect.Top, Txt��ˮ��.Height, blncancle, , True)
            If (Not rsTemp Is Nothing) And Not blncancle Then
                Call ���ﹺҩ��ϸ��ѯ(rsTemp!֧��˳���, rsTemp!����, rsTemp!�շ�ʱ��)
            End If
        Case "סԺ��ˮ��"
            'סԺȡסԺ�ǼǺ�
            gstrSQL = "select a.����id as id,B.����,A.˳���,to_char(A.����ʱ��,'yyyy-mm-dd hh24:mi:ss') as �Ǽ�ʱ��" & _
                      "  from �����ʻ� A,������Ϣ B where  A.����=" & TYPE_��ͨ & " and A.����id=" & Txt��ˮ��.Text & " and A.����id=B.����id  and nvl(b.סԺ����,0)>0 and A.����ʱ��>=to_date('" & Format(zlDatabase.Currentdate - mBusinessedDay, "yyyy-mm-dd HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')"
            Set rsTemp = zlDatabase.ShowSelect(Me, gstrSQL, 0, , , , , , , True, vRect.Left, vRect.Top, Txt��ˮ��.Height, blncancle, True, True)
            If (Not rsTemp Is Nothing) And Not blncancle Then
                Call סԺ�����ѯ(rsTemp!˳���, rsTemp!����, rsTemp!�Ǽ�ʱ��)
            End If
        End Select
    End If
End Sub
Private Function GetControlRect(ByVal lngHwnd As Long) As RECT
'���ܣ���ȡָ���ؼ�����Ļ�е�λ��(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function
Private Sub ���ﹺҩ��ϸ��ѯ(ByVal str֧��˳��� As String, ByVal str���� As String, ByVal str�շ�ʱ�� As String)
    '���ﹺҩ��ϸ��ѯ
Dim lngLoop As Long, strTemp As String
    
On Error GoTo errHandle
    
    Call InitTable
    If Not frmConn��ͨ.Execute("I290", 1, str֧��˳���, "���ڻ�ȡ���ﹺҩ��ϸ����......") Then Exit Sub
    Call ShowWindow(frmConn��ͨ.hwnd, 9)
    DoEvents
    For lngLoop = 1 To frmConn��ͨ.mlngRows
    DoEvents
        '��������
        If frmConn��ͨ.Query(lngLoop - 1, 1, "���ڲ�ѯ����(" & lngLoop & "/" & frmConn��ͨ.mlngRows & ")......") = False Then Exit Sub
        strTemp = frmConn��ͨ.strReturnInfo
        '��ʾ����
        With msh��ϸ_S
            If lngLoop > 1 Then .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = str֧��˳���
            .TextMatrix(.Rows - 1, 1) = Split(strTemp, vbTab)(0)
            .TextMatrix(.Rows - 1, 2) = Split(strTemp, vbTab)(1)
            .TextMatrix(.Rows - 1, 3) = Split(strTemp, vbTab)(2)
            .TextMatrix(.Rows - 1, 4) = Split(strTemp, vbTab)(3)
            .TextMatrix(.Rows - 1, 5) = Split(strTemp, vbTab)(4)
            .TextMatrix(.Rows - 1, 6) = Split(strTemp, vbTab)(5)
            .TextMatrix(.Rows - 1, 7) = Split(strTemp, vbTab)(6)
        End With
    Next
    Call ShowWindow(frmConn��ͨ.hwnd, 0)
    With msh��ϸ_S
        .Rows = .Rows + 2
        .TextMatrix(.Rows - 1, 0) = "����"
        .TextMatrix(.Rows - 1, 1) = str����
        .TextMatrix(.Rows - 1, 2) = "HIS�շ�ʱ��"
        .TextMatrix(.Rows - 1, 3) = str�շ�ʱ��
    End With
    Exit Sub
    
errHandle:
    If MsgBox("��ѯ����ʱ��������" & vbCrLf & Err.Description & vbCrLf & "�Ƿ����ԣ�", vbInformation + vbRetryCancel, "����") = vbRetry Then
        Err.Clear
        Resume
    End If
    Call ShowWindow(frmConn��ͨ.hwnd, 0)
End Sub
Private Sub סԺ�����ѯ(ByVal str˳��� As String, ByVal str���� As String, ByVal str�Ǽ�ʱ�� As String)
Dim lngLoop As Long, lngסԺ��¼ As Long, lngסԺ��ϸ As Long, strTemp As String
    
On Error GoTo errHandle
    
    Call InitTable
    '��ʼ��סԺ��¼��
    msh��ϸ_S.TextMatrix(1, 6) = "סԺ�ǼǼ�¼"
    
    If Not frmConn��ͨ.Execute("I360", 0, str˳���, "���ڻ�ȡסԺ���......") Then Exit Sub
    Call ShowWindow(frmConn��ͨ.hwnd, 9)
    DoEvents
    For lngLoop = 1 To frmConn��ͨ.mlngRows
        DoEvents
        '��������(סԺ��¼)
        If frmConn��ͨ.Query(lngLoop - 1, 1, "���ڲ�ѯ����(" & lngLoop & "/" & frmConn��ͨ.mlngRows & ")......") = False Then Exit Sub
        strTemp = frmConn��ͨ.strReturnInfo
        '��ʾ����
        With msh��ϸ_S
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = "סԺ���"
            .TextMatrix(.Rows - 1, 1) = "״̬"
            .TextMatrix(.Rows - 1, 2) = "����֤��"
            .TextMatrix(.Rows - 1, 3) = "����"
            .TextMatrix(.Rows - 1, 4) = "���"
            .TextMatrix(.Rows - 1, 5) = "�Ʊ�"
            .TextMatrix(.Rows - 1, 6) = "����"
            .TextMatrix(.Rows - 1, 7) = "��Ժ����"
            .TextMatrix(.Rows - 1, 8) = "��Ժ����"
            .TextMatrix(.Rows - 1, 9) = "סԺ����"
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = Split(strTemp, vbTab)(0)
            .TextMatrix(.Rows - 1, 1) = Split(strTemp, vbTab)(1)
            .TextMatrix(.Rows - 1, 2) = Split(strTemp, vbTab)(2)
            .TextMatrix(.Rows - 1, 3) = Split(strTemp, vbTab)(3)
            .TextMatrix(.Rows - 1, 4) = Split(strTemp, vbTab)(4)
            .TextMatrix(.Rows - 1, 5) = Split(strTemp, vbTab)(5)
            .TextMatrix(.Rows - 1, 6) = Split(strTemp, vbTab)(6)
            .TextMatrix(.Rows - 1, 7) = Split(strTemp, vbTab)(7)
            .TextMatrix(.Rows - 1, 8) = Split(strTemp, vbTab)(8)
            .TextMatrix(.Rows - 1, 9) = Split(strTemp, vbTab)(9)
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = "Ѻ��"
            .TextMatrix(.Rows - 1, 1) = "�𸶽�"
            .TextMatrix(.Rows - 1, 2) = "�޶�"
            .TextMatrix(.Rows - 1, 3) = "ҽ�Ʒ���"
            .TextMatrix(.Rows - 1, 4) = "ͳ�︺��"
            .TextMatrix(.Rows - 1, 5) = "���˸���"
            .TextMatrix(.Rows - 1, 6) = "�˻�֧��"
            .TextMatrix(.Rows - 1, 7) = "�����"
            .TextMatrix(.Rows - 1, 8) = "��������"
            .TextMatrix(.Rows - 1, 9) = "������"
            .TextMatrix(.Rows - 1, 10) = "����"
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = Split(strTemp, vbTab)(10)
            .TextMatrix(.Rows - 1, 1) = Split(strTemp, vbTab)(11)
            .TextMatrix(.Rows - 1, 2) = Split(strTemp, vbTab)(12)
            .TextMatrix(.Rows - 1, 3) = Split(strTemp, vbTab)(13)
            .TextMatrix(.Rows - 1, 4) = Split(strTemp, vbTab)(14)
            .TextMatrix(.Rows - 1, 5) = Split(strTemp, vbTab)(15)
            .TextMatrix(.Rows - 1, 6) = Split(strTemp, vbTab)(16)
            .TextMatrix(.Rows - 1, 7) = Split(strTemp, vbTab)(17)
            .TextMatrix(.Rows - 1, 8) = Split(strTemp, vbTab)(18)
            .TextMatrix(.Rows - 1, 9) = Split(strTemp, vbTab)(19)
            .TextMatrix(.Rows - 1, 10) = Split(strTemp, vbTab)(20)
        End With
    Next
    Call ShowWindow(frmConn��ͨ.hwnd, 0)
    
    
    '��ʼ��סԺ������ϸ��
    With msh��ϸ_S
        .Rows = .Rows + 2
        .TextMatrix(.Rows - 1, 6) = "סԺ��ϸ��¼"
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = "��ϸ���"
        .TextMatrix(.Rows - 1, 1) = "�ύ����"
        .TextMatrix(.Rows - 1, 2) = "����"
        .TextMatrix(.Rows - 1, 3) = "����"
        .TextMatrix(.Rows - 1, 4) = "����"
        .TextMatrix(.Rows - 1, 5) = "״̬"
        .TextMatrix(.Rows - 1, 6) = "��λ"
        .TextMatrix(.Rows - 1, 7) = "���"
        .TextMatrix(.Rows - 1, 8) = "����"
        .TextMatrix(.Rows - 1, 9) = "����"
        .TextMatrix(.Rows - 1, 10) = "���"
        .TextMatrix(.Rows - 1, 11) = "�����־"
    End With
    strTemp = str˳��� & vbTab & "0" & vbTab & "0" & vbTab & " " & vbTab & " " & vbTab & " "
    If Not frmConn��ͨ.Execute("I365", 0, strTemp, "���ڻ�ȡסԺ��ϸ��¼......") Then Exit Sub
    Call ShowWindow(frmConn��ͨ.hwnd, 9)
    DoEvents
    
    For lngLoop = 1 To frmConn��ͨ.mlngRows
        DoEvents
        '��������(סԺ������ϸ)
        If frmConn��ͨ.Query(lngLoop - 1, 1, "���ڲ�ѯ����(" & lngLoop & "/" & frmConn��ͨ.mlngRows & ")......") = False Then Exit Sub
        strTemp = frmConn��ͨ.strReturnInfo
        With msh��ϸ_S
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = Split(strTemp, vbTab)(0)
            .TextMatrix(.Rows - 1, 1) = Split(strTemp, vbTab)(1)
            .TextMatrix(.Rows - 1, 2) = Split(strTemp, vbTab)(2)
            .TextMatrix(.Rows - 1, 3) = Split(strTemp, vbTab)(3)
            .TextMatrix(.Rows - 1, 4) = Split(strTemp, vbTab)(4)
            .TextMatrix(.Rows - 1, 5) = Split(strTemp, vbTab)(5)
            .TextMatrix(.Rows - 1, 6) = Split(strTemp, vbTab)(6)
            .TextMatrix(.Rows - 1, 7) = Split(strTemp, vbTab)(7)
            .TextMatrix(.Rows - 1, 8) = Split(strTemp, vbTab)(8)
            .TextMatrix(.Rows - 1, 9) = Split(strTemp, vbTab)(9)
            .TextMatrix(.Rows - 1, 10) = Split(strTemp, vbTab)(10)
            .TextMatrix(.Rows - 1, 11) = Split(strTemp, vbTab)(11)
        End With
    Next
    lngסԺ��ϸ = lngLoop - 1
    Call ShowWindow(frmConn��ͨ.hwnd, 0)

    '��ʼ��סԺ���û�����
    With msh��ϸ_S
        .Rows = .Rows + 2
        .TextMatrix(.Rows - 1, 0) = "���û���"
        .TextMatrix(.Rows - 1, 2) = "����ͳ��:"
    End With
    strTemp = str˳���
    If Not frmConn��ͨ.Execute("I361", 5, strTemp, "���ڻ�ȡסԺ��ϸ��¼......") Then Exit Sub
    Call ShowWindow(frmConn��ͨ.hwnd, 9)
    DoEvents
    For lngLoop = 1 To frmConn��ͨ.mlngRows
        DoEvents
        '��������(סԺ������ϸ)
        If frmConn��ͨ.Query(lngLoop - 1, 1, "���ڲ�ѯ����(" & lngLoop & "/" & frmConn��ͨ.mlngRows & ")......") = False Then Exit Sub
        strTemp = frmConn��ͨ.strReturnInfo
        With msh��ϸ_S
            .TextMatrix(.Rows - 1, 1) = Split(strTemp, vbTab)(0)
            .TextMatrix(.Rows - 1, 3) = Split(strTemp, vbTab)(1)
        End With
    Next
    With msh��ϸ_S
        .Rows = .Rows + 2
        .TextMatrix(.Rows - 1, 0) = "����"
        .TextMatrix(.Rows - 1, 1) = str����
        .TextMatrix(.Rows - 1, 8) = "HIS�Ǽ�ʱ��"
        .TextMatrix(.Rows - 1, 10) = str�Ǽ�ʱ��
    End With
    Call ShowWindow(frmConn��ͨ.hwnd, 0)
    Exit Sub
    
errHandle:
    If MsgBox("��ѯ����ʱ��������" & vbCrLf & Err.Description & vbCrLf & "�Ƿ����ԣ�", vbInformation + vbRetryCancel, "����") = vbRetry Then
        Err.Clear
        Resume
    End If
    Call ShowWindow(frmConn��ͨ.hwnd, 0)
End Sub
