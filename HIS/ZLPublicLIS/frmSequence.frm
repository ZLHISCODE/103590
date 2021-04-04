VERSION 5.00
Begin VB.Form frmSequence 
   BorderStyle     =   0  'None
   ClientHeight    =   6210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.OptionButton opt类别 
      Caption         =   "仅名称"
      Height          =   255
      Index           =   3
      Left            =   4920
      TabIndex        =   13
      Top             =   180
      Width           =   1575
   End
   Begin VB.ListBox lst_Name 
      Appearance      =   0  'Flat
      Height          =   5490
      Left            =   3580
      Style           =   1  'Checkbox
      TabIndex        =   11
      Top             =   720
      Width           =   3210
   End
   Begin VB.ListBox lst_Type 
      Appearance      =   0  'Flat
      Height          =   5490
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   8
      Top             =   720
      Width           =   1550
   End
   Begin VB.ListBox lst_Sample 
      Appearance      =   0  'Flat
      Height          =   5490
      Left            =   1560
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   720
      Width           =   2000
   End
   Begin VB.OptionButton opt类别 
      Caption         =   "类别、标本混合"
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   6
      Top             =   180
      Width           =   1575
   End
   Begin VB.OptionButton opt类别 
      Caption         =   "仅标本"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   5
      Top             =   180
      Width           =   855
   End
   Begin VB.OptionButton opt类别 
      Caption         =   "仅类别"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   4
      Top             =   180
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton cmd取消 
      Caption         =   "取消"
      Height          =   300
      Left            =   6885
      TabIndex        =   3
      Top             =   2880
      Width           =   700
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定"
      Height          =   300
      Left            =   6885
      TabIndex        =   2
      Top             =   2280
      Width           =   700
   End
   Begin VB.CommandButton cmdFavorite 
      Caption         =   "下移"
      Height          =   300
      Index           =   1
      Left            =   6885
      TabIndex        =   1
      Top             =   1680
      Width           =   700
   End
   Begin VB.CommandButton cmdFavorite 
      Caption         =   "上移"
      Height          =   300
      Index           =   0
      Left            =   6885
      Picture         =   "frmSequence.frx":0000
      TabIndex        =   0
      Top             =   1080
      Width           =   700
   End
   Begin VB.Label lbl类别 
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      Caption         =   "名称"
      Height          =   180
      Index           =   2
      Left            =   3600
      TabIndex        =   12
      Top             =   480
      Width           =   360
   End
   Begin VB.Label lbl类别 
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      Caption         =   "类别"
      Height          =   180
      Index           =   0
      Left            =   0
      TabIndex        =   10
      Top             =   480
      Width           =   360
   End
   Begin VB.Label lbl类别 
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      Caption         =   "标本"
      Height          =   180
      Index           =   1
      Left            =   1560
      TabIndex        =   9
      Top             =   480
      Width           =   360
   End
End
Attribute VB_Name = "frmSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mType       As Integer

Private Sub cmdFavorite_Click(index As Integer)
    Dim strName     As String
    Dim blnSelect   As Boolean
    Dim objListBox  As ListBox
    
    If mType = 1 Then
        Set objListBox = lst_Type
    ElseIf mType = 2 Then
        Set objListBox = lst_Sample
    Else
        Set objListBox = lst_Name
    End If
    
    Select Case index
    Case 0              'up
        
        With objListBox
            If .ListIndex > 0 Then
                strName = .List(.ListIndex)
                blnSelect = .Selected(.ListIndex)
                
                .List(.ListIndex) = .List(.ListIndex - 1)
                .Selected(.ListIndex) = .Selected(.ListIndex - 1)
                
                .ListIndex = .ListIndex - 1
                
                .List(.ListIndex) = strName
                .Selected(.ListIndex) = blnSelect
                DataChanged = True
            End If
        End With
        
    Case 1              'down
    
        With objListBox
            If .ListIndex < .ListCount - 1 Then
                
                strName = .List(.ListIndex)
                blnSelect = .Selected(.ListIndex)
                
                .List(.ListIndex) = .List(.ListIndex + 1)
                .Selected(.ListIndex) = .Selected(.ListIndex + 1)
                
                .ListIndex = .ListIndex + 1
                
                .List(.ListIndex) = strName
                .Selected(.ListIndex) = blnSelect
                DataChanged = True
                
            End If
        End With
        
    End Select
End Sub

Private Sub cmd取消_Click()
    Unload Me
End Sub

Private Sub cmd确定_Click()
    Dim iCount  As Integer
    Dim iItem   As Integer
    Dim iType   As Integer
    Dim strSql  As String
    
    If opt类别(0).Value = True Then
        iType = 0
    End If
    If opt类别(1).Value = True Then
        iType = 1
    End If
    If opt类别(2).Value = True Then
        iType = 2
    End If
    If opt类别(3).Value = True Then
        iType = 3
    End If
    If iType = 0 Then
        For iItem = 0 To lst_Type.ListCount - 1
            strSql = "zl_类别顺序_Modify('" & lst_Type.List(iItem) & "'," & iItem + 1 & ")"
            Call gobjDatabase.ExecuteProcedure(strSql, gstrSysName)
        Next
        For iItem = 0 To lst_Sample.ListCount - 1
            strSql = "zl_标本顺序_Modify('" & lst_Sample.List(iItem) & "'," & 999 & ")"
            Call gobjDatabase.ExecuteProcedure(strSql, gstrSysName)
        Next
        For iItem = 0 To lst_Name.ListCount - 1
            strSql = "zl_项目顺序_Modify('" & lst_Name.List(iItem) & "'," & 999 & ")"
            Call gobjDatabase.ExecuteProcedure(strSql, gstrSysName)
        Next
    ElseIf iType = 1 Then
        For iItem = 0 To lst_Sample.ListCount - 1
            strSql = "zl_标本顺序_Modify('" & lst_Sample.List(iItem) & "'," & iItem + 1 & ")"
            Call gobjDatabase.ExecuteProcedure(strSql, gstrSysName)
        Next
        For iItem = 0 To lst_Type.ListCount - 1
            strSql = "zl_类别顺序_Modify('" & lst_Type.List(iItem) & "'," & 999 & ")"
            Call gobjDatabase.ExecuteProcedure(strSql, gstrSysName)
        Next
        For iItem = 0 To lst_Name.ListCount - 1
            strSql = "zl_项目顺序_Modify('" & lst_Name.List(iItem) & "'," & 999 & ")"
            Call gobjDatabase.ExecuteProcedure(strSql, gstrSysName)
        Next
    ElseIf iType = 2 Then
        For iItem = 0 To lst_Type.ListCount - 1
            strSql = "zl_类别顺序_Modify('" & lst_Type.List(iItem) & "'," & iItem + 1 & ")"
            Call gobjDatabase.ExecuteProcedure(strSql, gstrSysName)
        Next
        For iItem = 0 To lst_Sample.ListCount - 1
            strSql = "zl_标本顺序_Modify('" & lst_Sample.List(iItem) & "'," & iItem + 1 & ")"
            Call gobjDatabase.ExecuteProcedure(strSql, gstrSysName)
        Next
        For iItem = 0 To lst_Name.ListCount - 1
            strSql = "zl_项目顺序_Modify('" & lst_Name.List(iItem) & "'," & 999 & ")"
            Call gobjDatabase.ExecuteProcedure(strSql, gstrSysName)
        Next
    Else
        For iItem = 0 To lst_Type.ListCount - 1
            strSql = "zl_类别顺序_Modify('" & lst_Type.List(iItem) & "'," & 999 & ")"
            Call gobjDatabase.ExecuteProcedure(strSql, gstrSysName)
        Next
        For iItem = 0 To lst_Sample.ListCount - 1
            strSql = "zl_标本顺序_Modify('" & lst_Sample.List(iItem) & "'," & 999 & ")"
            Call gobjDatabase.ExecuteProcedure(strSql, gstrSysName)
        Next
        For iItem = 0 To lst_Name.ListCount - 1
            strSql = "zl_项目顺序_Modify('" & lst_Name.List(iItem) & "'," & iItem + 1 & ")"
            Call gobjDatabase.ExecuteProcedure(strSql, gstrSysName)
        Next
    End If
    
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    Dim index As Integer
    If gbtyModel = 1 Then
        Me.Top = 1560 + 1860
        Me.Left = 1560 + 5580
    Else
        Me.Top = glngTop + 545 ' 1560 + 1860
        Me.Left = glngLeft + 2300 ' 1560 + 5580
    End If
    index = JustType()

    Call opt类别_Click(index)
    Call LoadData
    If index = 0 Then
        
    ElseIf index = 1 Then
    ElseIf index = 2 Then
    Else
    End If
End Sub

Public Sub ShowMe(ByVal bytMode As Integer)
    Me.Show bytMode
End Sub

Private Property Let DataChanged(ByVal blnData As Boolean)
    cmd确定.Tag = IIf(blnData, "Changed", "")
End Property

Private Property Get DataChanged() As Boolean
    DataChanged = (cmd确定.Tag = "Changed")
End Property

Private Sub LoadData()
    Dim strSql  As String
    Dim rsType  As ADODB.Recordset
    Dim rsSamp  As ADODB.Recordset
    Dim rsName  As ADODB.Recordset
    
    strSql = ""
    
    Set rsType = GetLisType()
    Set rsSamp = GetLisSample()
    Set rsName = GetLisName()
    
    rsType.Filter = ""
    rsType.Sort = "顺序"
    
    rsSamp.Filter = ""
    rsSamp.Sort = "顺序"
    
    lst_Type.Clear
    lst_Sample.Clear
    
    If Not ChkRsState(rsSamp) Then
        rsSamp.MoveFirst
        Do While Not rsSamp.EOF
            lst_Sample.AddItem rsSamp("名称").Value
'            lstDept.ItemData(rsType.NewIndex) = rsType("ID").Value
            
            lst_Sample.Selected(lst_Sample.NewIndex) = True
            
            rsSamp.MoveNext
        Loop
    End If
    
    If Not ChkRsState(rsType) Then
        rsType.MoveFirst
        Do While Not rsType.EOF
            lst_Type.AddItem rsType("名称").Value
'            lstDept.ItemData(rsType.NewIndex) = rsType("ID").Value
            
            lst_Type.Selected(lst_Type.NewIndex) = True
            
            rsType.MoveNext
        Loop
    End If
    
    If Not ChkRsState(rsName) Then
        rsName.MoveFirst
        Do While Not rsName.EOF
            lst_Name.AddItem rsName("名称").Value
'            lstDept.ItemData(rsType.NewIndex) = rsType("ID").Value
            
            lst_Name.Selected(lst_Name.NewIndex) = True
            
            rsName.MoveNext
        Loop
    End If
    lst_Type.ListIndex = 0
    lst_Sample.ListIndex = 0
    lst_Name.ListIndex = 0
End Sub

Private Sub lst_Name_Click()
    mType = 3
End Sub

Private Sub lst_Sample_Click()
    mType = 2
End Sub

Private Sub lst_Type_Click()
    mType = 1
End Sub

Private Sub opt类别_Click(index As Integer)
    
    opt类别(index).Value = True
    If index = 0 Then
        lst_Type.Enabled = True
        lst_Sample.Enabled = False
        lst_Name.Enabled = False
        Call lst_Type_Click
    ElseIf index = 1 Then
        lst_Type.Enabled = False
        lst_Sample.Enabled = True
        lst_Name.Enabled = False
        Call lst_Sample_Click
    ElseIf index = 2 Then
        lst_Type.Enabled = True
        lst_Sample.Enabled = True
        lst_Name.Enabled = False
        Call lst_Type_Click
    Else
        lst_Type.Enabled = False
        lst_Sample.Enabled = False
        lst_Name.Enabled = True
        Call lst_Name_Click
    End If
End Sub
