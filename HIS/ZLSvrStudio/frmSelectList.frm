VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelectList 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3855
   ClientLeft      =   2730
   ClientTop       =   3435
   ClientWidth     =   5775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5775
   Begin VB.PictureBox picDrag 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   5115
      MousePointer    =   8  'Size NW SE
      Picture         =   "frmSelectList.frx":0000
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   1
      Top             =   2805
      Width           =   165
   End
   Begin MSComctlLib.StatusBar stb 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   3570
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10134
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2730
      Left            =   225
      TabIndex        =   0
      Top             =   615
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   4815
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.PictureBox picTitle 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   1935
      TabIndex        =   2
      Top             =   0
      Width           =   1935
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1380
         TabIndex        =   3
         Top             =   45
         Width           =   225
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H8000000E&
         Height          =   180
         Left            =   45
         TabIndex        =   5
         Top             =   45
         Width           =   360
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   4710
      Top             =   735
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectList.frx":0182
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSelectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnStartUp As Boolean
Private mlngIndex As Long
Private mblnOK As Boolean
Private mstrStatePath As String
Private mlngSortColumn As Long
Private mstrTitle As String
Private mlngX As Long
Private mlngY As Long

Private Sub SaveFormState()
    
    '���ܣ����浱ǰѡ������״̬
    
    Dim lngLoop As Long
    Dim strTmp As String
    
    If mstrStatePath = "" Then Exit Sub
    
    SaveSetting "ZLSOFT", mstrStatePath, "���", Me.Width
    SaveSetting "ZLSOFT", mstrStatePath, "�߶�", Me.Height
    
    For lngLoop = 1 To lvw.ColumnHeaders.Count
        strTmp = strTmp & ";" & lvw.ColumnHeaders(lngLoop).Width
    Next
    If strTmp <> "" Then strTmp = Mid(strTmp, 2)
    
    SaveSetting "ZLSOFT", mstrStatePath, "�п�", strTmp
    
End Sub

Private Sub RestoreFormState()
    
    '���ܣ����浱ǰѡ������״̬
    
    Dim lngLoop As Long
    Dim strTmp As String
    
    If mstrStatePath = "" Then Exit Sub
    
    Me.Width = GetSetting("ZLSOFT", mstrStatePath, "���", Me.Width)
    Me.Height = GetSetting("ZLSOFT", mstrStatePath, "�߶�", Me.Height)
    
    For lngLoop = 1 To lvw.ColumnHeaders.Count
        strTmp = strTmp & ";" & lvw.ColumnHeaders(lngLoop).Width
    Next
    If strTmp <> "" Then strTmp = Mid(strTmp, 2)
    
    strTmp = GetSetting("ZLSOFT", mstrStatePath, "�п�", strTmp)
    
    On Error Resume Next
    
    For lngLoop = 1 To lvw.ColumnHeaders.Count
        lvw.ColumnHeaders(lngLoop).Width = Val(Split(strTmp, ";")(lngLoop - 1))
    Next
    
End Sub

Public Function ShowSelect(ByVal frmMain As Form, _
                            ByRef rsData As ADODB.Recordset, _
                            ByVal strLvw As String, _
                            ByVal sglX As Single, _
                            ByVal sglY As Single, _
                            ByVal sglCX As Single, _
                            ByVal sglCY As Single, _
                            Optional ByVal StatePath As String, _
                            Optional strTitle As String, _
                            Optional BackColor As Long = &H80000005, _
                            Optional InitSelectKey As String = "", _
                            Optional notblnRestoryWindow As Boolean = False _
                            ) As Boolean
    
    '����:��ʾ��ѯѡ����
    '����:
    '����:
    
    Dim lngLoop As Long
    Dim objItem As ListItem
    
    If rsData.BOF Then Exit Function
    
    rsData.MoveFirst
    
    mblnStartUp = True
    mblnOK = False
    mstrStatePath = "˽��ģ��\" & gstrUserName & "\" & App.ProductName & "\" & StatePath
    mlngSortColumn = 1
    mstrTitle = strTitle
        
    lvw.BackColor = BackColor
    lvw.ListItems.Clear
    LvwSelectColumns lvw, strLvw, True
        
    Me.Left = sglX
    Me.Top = sglY
    Me.Width = sglCX
    Me.Height = sglCY
    lblCaption.Caption = strTitle
                    
    '�ָ������С���б���
    If notblnRestoryWindow = False Then
        Call RestoreFormState
    End If
    'װ������
    With lvw
        Do While Not rsData.EOF
            
            Set objItem = .ListItems.Add(, , Nvl(rsData(.ColumnHeaders(1).Text), ""), 1, 1)
            For lngLoop = 1 To .ColumnHeaders.Count - 1
                objItem.SubItems(lngLoop) = IIf(IsNull(rsData(.ColumnHeaders(lngLoop + 1).Text)), "", rsData(.ColumnHeaders(lngLoop + 1).Text))
            Next
            rsData.MoveNext
        Loop
    End With
    
    stb.Panels(1).Text = "�������� " & lvw.ListItems.Count & " �������"
    
    On Error Resume Next
    If InitSelectKey <> "" Then
        
        lvw.ListItems("K" & InitSelectKey).Selected = True
        lvw.ListItems("K" & InitSelectKey).EnsureVisible
        
    End If
    On Error GoTo 0
    
    Me.Show 1, frmMain
    
    rsData.MoveFirst
    rsData.Move mlngIndex - 1
    
    ShowSelect = mblnOK
    
End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub ReturnSelect()
    mlngIndex = lvw.SelectedItem.Index
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With picTitle
        .Left = -15
        .Top = -30
        .Width = Me.ScaleWidth + 30
    End With
    
    With lvw
        .Left = -15
        .Top = picTitle.Top + picTitle.Height
        .Width = Me.ScaleWidth + 30
        .Height = Me.ScaleHeight - stb.Height - .Top
    End With
    
    With picDrag
        .Left = Me.ScaleWidth - .Width - 30
        .Top = Me.ScaleHeight - .Height - 30
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveFormState
End Sub

Private Sub lvw_DblClick()
    Call ReturnSelect
End Sub

Private Sub lvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call lvw_DblClick
End Sub

Private Sub picDrag_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Err = 0: On Error Resume Next
    If Button = 1 Then
        mlngX = x
        mlngY = y
    End If
End Sub

Private Sub picDrag_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Err = 0: On Error Resume Next
    If Button = 1 Then
        Me.Width = Me.Width + x - mlngX
        Me.Height = Me.Height + y - mlngY
        Call Form_Resize
    End If
End Sub

Private Sub picTitle_Resize()
    On Error Resume Next
    
    With cmdClose
        .Left = picTitle.Width - .Width - 30
    End With
End Sub
