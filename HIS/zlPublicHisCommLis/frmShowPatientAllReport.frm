VERSION 5.00
Begin VB.Form frmShowPatientAllReport 
   Caption         =   "打印预览"
   ClientHeight    =   9090
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12390
   Icon            =   "frmShowPatientAllReport.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   9090
   ScaleWidth      =   12390
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6915
      Left            =   -90
      ScaleHeight     =   6885
      ScaleWidth      =   13785
      TabIndex        =   0
      Top             =   810
      Width           =   13815
      Begin VB.Frame fraWE 
         BorderStyle     =   0  'None
         Height          =   5115
         Left            =   5820
         MousePointer    =   9  'Size W E
         TabIndex        =   3
         Top             =   870
         Width           =   105
      End
      Begin VB.PictureBox picNewReport 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5235
         Left            =   810
         ScaleHeight     =   5205
         ScaleWidth      =   3705
         TabIndex        =   2
         Top             =   750
         Width           =   3735
      End
      Begin VB.PictureBox picOldReport 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5055
         Left            =   7740
         ScaleHeight     =   5025
         ScaleWidth      =   4185
         TabIndex        =   1
         Top             =   1110
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmShowPatientAllReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'将控件放到容器中
Private Declare Function SetParent Lib "user32.dll " (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_MAXIMIZE = &HF030&
Private Const SC_RESTORE = &HF120&

Private mobjNewReport As Object
Private mobjOldReport As Object

Private Sub Form_Resize()
    On Error Resume Next
    With Me.picMain
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjNewReport = Nothing
    Set mobjOldReport = Nothing
End Sub

Private Sub fraWE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim LeftColl As New Collection, Rightcoll As New Collection
    If Button = vbLeftButton Then
        LeftColl.Add Me.picNewReport
        Rightcoll.Add Me.picOldReport
        Call SplitWE(LeftColl, Me.fraWE, Rightcoll, X, 1000)
        Set LeftColl = Nothing
        Set Rightcoll = Nothing
    End If
End Sub

Private Sub PicMain_Resize()
    On Error Resume Next

    If Not mobjNewReport Is Nothing And Not mobjOldReport Is Nothing Then
        With Me.picNewReport
            .Left = 0
            .Top = 0
            .Width = (picMain.Width - Me.fraWE.Width) / 2
            .Height = Me.picMain.Height
            .Visible = True
        End With
        With Me.fraWE
            .Left = Me.picNewReport.Width
            .Top = 0
            .Height = Me.picMain.Height
            .Visible = True
        End With
        With Me.picOldReport
            .Left = Me.fraWE.Left + Me.fraWE.Width
            .Top = 0
            .Width = Me.picMain.Width - .Left
            .Height = Me.picMain.Height
            .Visible = True
        End With
    ElseIf Not mobjNewReport Is Nothing Then
        With picNewReport
            .Left = 0
            .Top = 0
            .Width = Me.picMain.Width
            .Height = Me.picMain.Height
            .Visible = True
        End With
        fraWE.Visible = False
        picOldReport.Visible = False
    ElseIf Not mobjOldReport Is Nothing Then
        With picOldReport
            .Left = 0
            .Top = 0
            .Width = Me.picMain.Width
            .Height = Me.picMain.Height
            .Visible = True
        End With
        fraWE.Visible = False
        picNewReport.Visible = False
    Else
        With picNewReport
            .Left = 0
            .Top = 0
            .Width = Me.picMain.Width
            .Height = Me.picMain.Height
        End With
        fraWE.Visible = False
        picOldReport.Visible = False
    End If
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-07-26
'功    能:  获取选中科室病人
'入    参:
'           lngPaitID           病人ID
'           intPage             主页ID
'出    参:
'返    回:
'调整影响:
'调用注意:
'---------------------------------------------------------------------------------------
Public Function ShowMe(objFrm As Object, ByVal lngPaitID As Long, ByVal intPage As Integer)
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strNewSampleIDs As String
          Dim strOldSampleIDs As String

          '新版报告
1         On Error GoTo showMe_Error

2         strSQL = "Select f_List2str(Cast(Collect(to_char(ID)) As t_Strlist)) 标本ID" & vbCrLf & _
                 "   From 检验报告记录" & vbCrLf & _
                 "   Where HIS病人ID = [1] and 审核人 is not null"
3         If intPage > 0 Then
4             strSQL = strSQL & " And 主页ID = [2]"
5         End If
6         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "新版报告", lngPaitID, intPage)
7         If Not rsTmp.EOF Then
8             strNewSampleIDs = rsTmp("标本ID") & ""
9         End If

          '老版报告
10        strSQL = "Select f_List2str(Cast(Collect(to_char(ID)) As t_Strlist)) 标本ID" & vbCrLf & _
                 "   From 检验标本记录" & vbCrLf & _
                 "   Where 病人ID = [1] and 审核人 is not null"
11        If intPage > 0 Then
12            strSQL = strSQL & " And 主页ID = [2]"
13        End If
14        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "新版报告", lngPaitID, intPage)
15        If Not rsTmp.EOF Then
16            strOldSampleIDs = rsTmp("标本ID") & ""
17        End If

          '加载报告
18        Call initReport
19        If Not mobjNewReport Is Nothing Then Unload mobjNewReport
20        If Not mobjOldReport Is Nothing Then Unload mobjOldReport
21        Set mobjNewReport = Nothing
22        Set mobjOldReport = Nothing
23        If strNewSampleIDs <> "" Then
24            Call zlReport.LoadReport(gcnLisOracle, 2500, "ZL25_INSIDE_2500_109", Me, mobjNewReport, Nothing, "标本ID=" & strNewSampleIDs, "DisabledPrint=1")
25        End If
26        If strOldSampleIDs <> "" Then
27            Call zlReport.LoadReport(gcnHisOracle, 100, "ZL1_INSIDE_1208_9", Me, mobjOldReport, Nothing, "标本ID=" & strOldSampleIDs, "DisabledPrint=1")
28        End If
29        Call PicMain_Resize
30        If Not mobjNewReport Is Nothing Then
31            Call LockWindowUpdate(mobjNewReport.hWnd)
32            SetParent mobjNewReport.hWnd, picNewReport.hWnd
33            Call SendMessage(mobjNewReport.hWnd, WM_SYSCOMMAND, SC_RESTORE, 0)
34            Call SendMessage(mobjNewReport.hWnd, WM_SYSCOMMAND, SC_MAXIMIZE, 0)
35            Call LockWindowUpdate(0)
36        End If
37        If Not mobjOldReport Is Nothing Then
38            Call LockWindowUpdate(mobjOldReport.hWnd)
39            SetParent mobjOldReport.hWnd, picOldReport.hWnd
40            Call SendMessage(mobjOldReport.hWnd, WM_SYSCOMMAND, SC_RESTORE, 0)
41            Call SendMessage(mobjOldReport.hWnd, WM_SYSCOMMAND, SC_MAXIMIZE, 0)
42            Call LockWindowUpdate(0)
43        End If

44        Me.Show vbModal, objFrm


45        Exit Function
showMe_Error:
46        Call WriteErrLog("zlPublicHisCommLis", "frmShowPatientAllReport", "执行(ShowMe)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
47        Err.Clear
End Function

Private Sub picNewReport_Resize()
    Dim H As Long, W As Long

    On Error Resume Next
     
    W = picNewReport.ScaleX(picNewReport.Width, picNewReport.ScaleMode, vbPixels)
    H = picNewReport.ScaleY(picNewReport.Height, picNewReport.ScaleMode, vbPixels)
    Call MoveWindow(mobjNewReport.hWnd, 0, 0, W, H, True)
End Sub

Private Sub picOldReport_Resize()
    Dim H As Long, W As Long

    On Error Resume Next

    W = picOldReport.ScaleX(picOldReport.Width, picOldReport.ScaleMode, vbPixels)
    H = picOldReport.ScaleY(picOldReport.Height, picOldReport.ScaleMode, vbPixels)
    Call MoveWindow(mobjOldReport.hWnd, 0, 0, W, H, True)
End Sub
