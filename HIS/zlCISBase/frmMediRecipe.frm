VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmMediRecipe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "??ҩ?䷽?༭"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10395
   Icon            =   "frmMediRecipe.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '??????????
   Begin VB.OptionButton optDosageType 
      Caption         =   "??????̬"
      Height          =   255
      Index           =   0
      Left            =   825
      TabIndex        =   43
      Top             =   2640
      Width           =   1095
   End
   Begin VB.OptionButton optDosageType 
      Caption         =   "??????"
      Height          =   255
      Index           =   3
      Left            =   5040
      TabIndex        =   42
      Top             =   2640
      Width           =   855
   End
   Begin VB.OptionButton optDosageType 
      Caption         =   "??Ƭ"
      Height          =   255
      Index           =   2
      Left            =   3755
      TabIndex        =   41
      Top             =   2640
      Width           =   735
   End
   Begin VB.OptionButton optDosageType 
      Caption         =   "ɢװ"
      Height          =   255
      Index           =   1
      Left            =   2470
      TabIndex        =   40
      Top             =   2640
      Width           =   735
   End
   Begin VB.ComboBox cmbStationNo 
      Height          =   300
      Left            =   825
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   2160
      Visible         =   0   'False
      Width           =   3600
   End
   Begin VB.CommandButton cmd?ο? 
      Caption         =   "??"
      Height          =   285
      Left            =   4140
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1800
      Width           =   285
   End
   Begin VB.TextBox txt?ο? 
      Height          =   300
      Left            =   825
      TabIndex        =   34
      Top             =   1785
      Width           =   3240
   End
   Begin VB.TextBox txt˵?? 
      Height          =   300
      Left            =   5820
      MaxLength       =   30
      TabIndex        =   17
      Top             =   1815
      Width           =   3600
   End
   Begin VB.TextBox txt???? 
      Height          =   300
      Index           =   1
      Left            =   825
      MaxLength       =   40
      TabIndex        =   12
      Top             =   1425
      Width           =   3600
   End
   Begin VB.TextBox txtƴ?? 
      Height          =   300
      Index           =   1
      Left            =   5820
      MaxLength       =   12
      TabIndex        =   14
      Top             =   1425
      Width           =   1350
   End
   Begin VB.TextBox txt???? 
      Height          =   300
      Index           =   1
      Left            =   7800
      MaxLength       =   12
      TabIndex        =   15
      Top             =   1425
      Width           =   960
   End
   Begin VB.TextBox txt?Ƴ? 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   825
      MaxLength       =   50
      TabIndex        =   26
      Top             =   5910
      Width           =   1020
   End
   Begin VB.ComboBox cboƵ?? 
      Height          =   300
      Left            =   780
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   5520
      Width           =   2115
   End
   Begin VB.ComboBox cbo?巨 
      Height          =   300
      Left            =   4155
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   5520
      Width           =   2115
   End
   Begin VB.ComboBox cbo?÷? 
      Height          =   300
      Left            =   7530
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   5520
      Width           =   2115
   End
   Begin VB.Frame fraLine 
      Height          =   15
      Index           =   1
      Left            =   -45
      TabIndex        =   33
      Top             =   6360
      Width           =   10410
   End
   Begin VB.TextBox txt???? 
      Height          =   300
      Index           =   0
      Left            =   7800
      MaxLength       =   12
      TabIndex        =   10
      Top             =   1050
      Width           =   960
   End
   Begin VB.TextBox txtƴ?? 
      Height          =   300
      Index           =   0
      Left            =   5820
      MaxLength       =   12
      TabIndex        =   9
      Top             =   1050
      Width           =   1350
   End
   Begin VB.TextBox txt???? 
      Height          =   300
      Index           =   0
      Left            =   825
      MaxLength       =   40
      TabIndex        =   7
      Top             =   1050
      Width           =   3600
   End
   Begin VB.TextBox txt???? 
      Height          =   300
      Left            =   825
      MaxLength       =   13
      TabIndex        =   2
      Top             =   675
      Width           =   3600
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2280
      Left            =   4080
      TabIndex        =   31
      Top             =   8400
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4022
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   3570
      Left            =   120
      TabIndex        =   30
      TabStop         =   0   'False
      Tag             =   "1000"
      Top             =   8400
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   6297
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgList"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ??(&C)"
      Height          =   350
      Left            =   8880
      TabIndex        =   28
      Top             =   6525
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "????(&H)"
      Height          =   350
      Left            =   135
      Picture         =   "frmMediRecipe.frx":058A
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   6525
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ??(&O)"
      Height          =   350
      Left            =   7800
      TabIndex        =   27
      Top             =   6525
      Width           =   1100
   End
   Begin VB.TextBox txt???? 
      Height          =   300
      Left            =   5820
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   4
      Top             =   675
      Width           =   3255
   End
   Begin VB.CommandButton cmd???? 
      Caption         =   "&P"
      Height          =   285
      Left            =   9105
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   690
      Width           =   285
   End
   Begin VB.Frame fraLine 
      Height          =   60
      Index           =   0
      Left            =   0
      TabIndex        =   32
      Top             =   480
      Width           =   10410
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   6960
      Top             =   8280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediRecipe.frx":06D4
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediRecipe.frx":0C6E
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediRecipe.frx":1208
            Key             =   "ҩƷ"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediRecipe.frx":17A2
            Key             =   "??ע"
         EndProperty
      EndProperty
   End
   Begin ZL9BillEdit.BillEdit msfRecipe 
      Height          =   2415
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   4260
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.Label lblDosageType 
      AutoSize        =   -1  'True
      Caption         =   "??̬(&B)"
      Height          =   180
      Left            =   150
      TabIndex        =   39
      Top             =   2640
      Width           =   630
   End
   Begin VB.Label lblStationNo 
      AutoSize        =   -1  'True
      Caption         =   "Ժ??(&Z)"
      Height          =   180
      Left            =   150
      TabIndex        =   38
      Top             =   2220
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "?ο?(&F)"
      Height          =   180
      Left            =   150
      TabIndex        =   36
      Top             =   1845
      Width           =   630
   End
   Begin VB.Label lbl˵?? 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "˵??(&Z)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5175
      TabIndex        =   16
      Top             =   1875
      Width           =   630
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   135
      Picture         =   "frmMediRecipe.frx":207C
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lbl???? 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "????(&A)"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   135
      TabIndex        =   11
      Top             =   1485
      Width           =   630
   End
   Begin VB.Label lbl???? 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "????(&M)                (ƴ??)            (????)"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   5160
      TabIndex        =   13
      Top             =   1485
      Width           =   4230
   End
   Begin VB.Label lbl?Ƴ? 
      AutoSize        =   -1  'True
      Caption         =   "?Ƴ?(&T)            (??????????Ҫ??һ????Ҫ???????ø??䷽?ĳ???????)??"
      Height          =   180
      Left            =   135
      TabIndex        =   25
      Top             =   5955
      Width           =   6210
   End
   Begin VB.Label lblƵ?? 
      AutoSize        =   -1  'True
      Caption         =   "Ƶ??(&L)"
      Height          =   180
      Left            =   135
      TabIndex        =   19
      Top             =   5580
      Width           =   630
   End
   Begin VB.Label lbl?巨 
      AutoSize        =   -1  'True
      Caption         =   "?巨(&J)"
      Height          =   180
      Left            =   3510
      TabIndex        =   21
      Top             =   5580
      Width           =   630
   End
   Begin VB.Label lbl?÷? 
      AutoSize        =   -1  'True
      Caption         =   "?÷?(&U)"
      Height          =   180
      Left            =   6885
      TabIndex        =   23
      Top             =   5580
      Width           =   630
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    ???鷽ԭ?򣬲ο?Ȩ?????????ϣ????в?ҩ???ɳ??õ??䷽???Է???ҽ???´???ҩҽ??ʱ????Ѹ??׼ȷ?????ɴ?????"
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   780
      TabIndex        =   0
      Top             =   105
      Width           =   9645
   End
   Begin VB.Label lbl???? 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "????(&S)                (ƴ??)            (????)"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   5160
      TabIndex        =   8
      Top             =   1110
      Width           =   4230
   End
   Begin VB.Label lbl???? 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "????(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   6
      Top             =   1110
      Width           =   630
   End
   Begin VB.Label lbl???? 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "????(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   135
      TabIndex        =   1
      Top             =   735
      Width           =   630
   End
   Begin VB.Label lbl???? 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "????(&F)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5160
      TabIndex        =   3
      Top             =   735
      Width           =   630
   End
End
Attribute VB_Name = "frmMediRecipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'˵????
'   1???ϼ?????ͨ????????ShowMe?????????????塢Ȩ?ޡ??༭??Ŀ?ķ???ID??ID,?༭״̬????Ϣ???ݽ??뱾????
'   2???༭״̬????Me.tag???ţ??ֱ?Ϊ"????"??"?޸?"??"????"?????ϼ?????ͨ??ShowMe????
'---------------------------------------------------
Private lngClassId As Long       '???༭?ķ???ID???ϼ?????ͨ??ShowMe???ݽ???
Private lngItemId As Long        '???༭????ĿID???޸ġ?????ʱ???ϼ?????ͨ??ShowMe???ݽ???,????ʱΪ0??
Private mbyt??ҩζ?? As Byte    '??ҩ?䷽ÿ????ҩζ??

Private strInputed As String
Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim objItem As ListItem
Dim strTemp As String, aryTemp() As String
Dim intCount As Integer, intFence As Integer
Private Const mstrChar As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789."
Private mblnDosage As Boolean '?Ƿ????????䷽
Private mstrMedi As String  '?༭֮ǰ

Dim mstrMatch As String, strRefer As String '?ο?????
Private mblnOK As Boolean
Private mintOldShape As Integer '??¼ѡ?е???̬ '0-??????̬ 1-ɢװ 2-??Ƭ 3-??????
Private mblnLoad As Boolean  '?????Ƿ??????? true-?????? false-δ??????
Private mblnClickNo As Boolean  '????????ʾ???еġ??񡱰?ť true-??????
Private Enum ?䷽?б?
    ?հ? = 0
    ҩ??ID = 1
    ????ID = 2
    ???? = 3
    ???? = 4
    ??λ = 5
    ??ע = 6
    ???? = 7
End Enum
Private Sub GetDefineSize()
    '???ܣ??õ????ݿ??ı??ֶεĳ???
    On Error GoTo errHandle
    Dim rsTmp As New ADODB.Recordset
    gstrSql = "Select A.????,A.?걾??λ,B.????,B.???? From ??????ĿĿ¼ A, ??????Ŀ???? B Where A.ID=B.??????ĿID and A.ID=0"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSql, Me.Caption)
    
    txt????.MaxLength = rsTmp.Fields("????").DefinedSize
    txt????(0).MaxLength = rsTmp.Fields("????").DefinedSize
    txt????(1).MaxLength = rsTmp.Fields("????").DefinedSize
    txtƴ??(0).MaxLength = rsTmp.Fields("????").DefinedSize
    txtƴ??(1).MaxLength = rsTmp.Fields("????").DefinedSize
    txt????(0).MaxLength = rsTmp.Fields("????").DefinedSize
    txt????(1).MaxLength = rsTmp.Fields("????").DefinedSize
    txt˵??.MaxLength = rsTmp.Fields("?걾??λ").DefinedSize

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Function ShowMe(ByVal frmParent As Object, ByVal byt״̬ As Byte, ByVal lng????id As Long, Optional ByVal lng??Ŀid As Long) As Boolean
    '---------------------------------------------------
    '???ܣ??ϼ????????ñ??????ģ????ݲ?????????ʾ????
    '---------------------------------------------------
    Dim intDosageType As Integer
    
    Me.Tag = Switch(byt״̬ = 0, "????", byt״̬ = 1, "?޸?", byt״̬ = 2, "????")
    lngClassId = lng????id: lngItemId = lng??Ŀid
    
    '??д??Ҫѡ????????
    Err = 0: On Error GoTo ErrHand
    
    If Me.Tag = "????" Then
        intDosageType = GetSetting("ZLSOFT", "˽??ģ??\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "?䷽??̬", 0)
    Else
        gstrSql = "select distinct nvl(?䷽????,0) ?䷽???? from ??????Ŀ???? where ????????ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "showme", lngItemId)
        If rsTemp.RecordCount > 0 Then
            intDosageType = rsTemp!?䷽????
        End If
    End If
    
    If intDosageType < 0 Or intDosageType > 3 Then
        intDosageType = 0
    End If
    optDosageType(intDosageType).Value = True
    
    gstrSql = "select ID,?ϼ?ID,????,????,????" & _
            " From ???Ʒ???Ŀ¼" & _
            " Where ???? = 4" & _
            " start with ?ϼ?ID is null" & _
            " connect by prior ID=?ϼ?ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ShowMe")

     With rsTemp
        If .BOF Or .EOF Then MsgBox "?????Ƚ????䷽???Ʒ?????Ŀ֮???????䷽", vbExclamation, gstrSysName: Unload Me: Exit Function
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
            If IsNull(!?ϼ?ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, "[" & !???? & "]" & !????, "close")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !?ϼ?ID, tvwChild, "_" & !ID, "[" & !???? & "]" & !????, "close")
            End If
            objNode.Sorted = True
            objNode.Tag = IIf(IsNull(!????), "", !????)
            objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
        Me.tvwClass.Nodes("_" & lng????id).Selected = True
        Me.txt????.Text = Me.tvwClass.SelectedItem.Text
        Me.txt????.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
        
        gstrSql = "select ????||'-'||???? as ???? from ????Ƶ????Ŀ where ???÷?Χ=2 order by ????"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ShowMe")
        
        Me.cboƵ??.Clear
        Do While Not rsTemp.EOF
            Me.cboƵ??.AddItem rsTemp!????
            rsTemp.MoveNext
        Loop
        If Me.cboƵ??.ListCount = 0 Then
            Me.cboƵ??.Enabled = False
        Else
            Me.cboƵ??.ListIndex = 0
        End If
        
        gstrSql = "select ID,rownum||'-'||???? as ???? from ??????ĿĿ¼ where ????='E' and ????????='3' order by ????"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ShowMe")

        Me.cbo?巨.Clear
        Me.cbo?巨.AddItem "": Me.cbo?巨.ItemData(Me.cbo?巨.NewIndex) = 0
        Do While Not rsTemp.EOF
            Me.cbo?巨.AddItem rsTemp!????: Me.cbo?巨.ItemData(Me.cbo?巨.NewIndex) = rsTemp!ID
            rsTemp.MoveNext
        Loop
        If Me.cbo?巨.ListCount = 0 Then
            Me.cbo?巨.Enabled = False
        Else
            Me.cbo?巨.ListIndex = 0
        End If
        
        gstrSql = "select ID,rownum||'-'||???? as ???? from ??????ĿĿ¼ where ????='E' and ????????='4' order by ????"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ShowMe")

        Me.cbo?÷?.Clear
        Do While Not rsTemp.EOF
            Me.cbo?÷?.AddItem rsTemp!????: Me.cbo?÷?.ItemData(Me.cbo?÷?.NewIndex) = rsTemp!ID
            rsTemp.MoveNext
        Loop
        If Me.cbo?÷?.ListCount = 0 Then
            Me.cbo?÷?.Enabled = False
        Else
            Me.cbo?÷?.ListIndex = 0
        End If
    End With
    If Me.cbo?÷?.Enabled = False And Me.cbo?巨.Enabled = False Then
        Me.cboƵ??.Enabled = False: Me.txt?Ƴ?.Enabled = False
    End If

    '??ʾ????
    Me.Show 1, frmParent
    ShowMe = mblnOK
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cbo?巨_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboƵ??_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo?÷?_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim strTmp As String
    Dim strSql As String
    Dim strվ?? As String
    Dim str???? As String
    Dim i As Integer
    Dim intDosage As Integer '?䷽????
    Dim intGroup As Integer '?ڼ???    һ????ζҩ?????? һ????λҩ??4??.....
    Dim strCheckID As String    'ҩ??id+????id
    
    If optDosageType(3).Value = True And cbo?巨.ItemData(cbo?巨.ListIndex) <> 0 Then
        MsgBox "?????????????ü巨??", vbInformation, gstrSysName
        cbo?巨.SetFocus
        Exit Sub
    End If
    
    For i = 0 To optDosageType.UBound
        If optDosageType(i).Value = True Then
            intDosage = i
            Exit For
        End If
    Next
    
    '???¼??????ƣ???ȥ???????ַ?
    strTmp = MoveSpecialChar(txt????(0).Text)
    If txt????(0).Text <> strTmp Then
        txt????(0).Text = strTmp
        Me.txtƴ??(0).Text = zlStr.GetCodeByORCL(Me.txt????(0).Text, False)
        Me.txt????(0).Text = zlStr.GetCodeByORCL(Me.txt????(0).Text, True)
    End If
    strTmp = MoveSpecialChar(txt????(1).Text)
    If txt????(1).Text <> strTmp Then
        txt????(1).Text = strTmp
        Me.txtƴ??(1).Text = zlStr.GetCodeByORCL(Me.txt????(1).Text, False)
        Me.txt????(1).Text = zlStr.GetCodeByORCL(Me.txt????(1).Text, True)
    End If
    
    'һ?????Լ???
    If Trim(Me.txt????.Text) = "" Then MsgBox "?????????룡", vbInformation, gstrSysName: Me.txt????.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt????.Text), vbFromUnicode)) > Me.txt????.MaxLength Then MsgBox "?????ĳ?????????" & Me.txt????.MaxLength & "???ַ?????", vbInformation, gstrSysName: Me.txt????.SetFocus: Exit Sub
    If Trim(Me.txt????(0).Text) = "" Then MsgBox "?????????ƣ?", vbInformation, gstrSysName: Me.txt????(0).SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt????(0).Text), vbFromUnicode)) > Me.txt????(0).MaxLength Then
        MsgBox "???Ƴ?????" & Me.txt????(0).MaxLength & "???ַ???" & Me.txt????(0).MaxLength / 2 & "?????֣???", vbInformation, gstrSysName: Me.txt????(0).SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt????(1).Text), vbFromUnicode)) > Me.txt????(1).MaxLength Then
        MsgBox "??????????" & Me.txt????(1).MaxLength & "???ַ???" & Me.txt????(1).MaxLength / 2 & "?????֣???", vbInformation, gstrSysName: Me.txt????(1).SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txtƴ??(0).Text), vbFromUnicode)) > Me.txtƴ??(0).MaxLength Then
        MsgBox "????ƴ?????볬????" & Me.txtƴ??(0).MaxLength & "???ַ?????", vbInformation, gstrSysName: Me.txtƴ??(0).SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txtƴ??(1).Text), vbFromUnicode)) > Me.txtƴ??(1).MaxLength Then
        MsgBox "????ƴ?????볬????" & Me.txtƴ??(1).MaxLength & "???ַ?????", vbInformation, gstrSysName: Me.txtƴ??(1).SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt????(0).Text), vbFromUnicode)) > Me.txt????(0).MaxLength Then
        MsgBox "???????ʼ??볬????" & Me.txt????(0).MaxLength & "???ַ?????", vbInformation, gstrSysName: Me.txt????(0).SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt????(1).Text), vbFromUnicode)) > Me.txt????(1).MaxLength Then
        MsgBox "???????ʼ??볬????" & Me.txt????(1).MaxLength & "???ַ?????", vbInformation, gstrSysName: Me.txt????(1).SetFocus: Exit Sub
    End If
    If Val(Me.txt?Ƴ?.Text) > 100 Then MsgBox "ϵͳ??????????̫?????Ƴ̣?Ϊ0??ʾ???????Ƴ̣???", vbExclamation, gstrSysName: Me.txt?Ƴ?.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt˵??.Text), vbFromUnicode)) > Me.txt˵??.MaxLength Then
        MsgBox "˵????????" & Me.txt˵??.MaxLength & "???ַ???" & Me.txt˵??.MaxLength / 2 & "?????֣???", vbInformation, gstrSysName: Me.txt˵??.SetFocus: Exit Sub
    End If
    
    '??????Ŀʱ????֤???????ظ????룬???????ظ??Զ???ԭ?????????ϼ?1??ֱ?????ظ?
    str???? = Trim(txt????.Text)
    If Me.Tag = "????" Then
        Do While True
            gstrSql = "select a.???? from ??????ĿĿ¼ a,??????Ŀ???? b where a.????=[1] and a.????=b.????"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "?????Ƿ??ظ?", str????)
            If rsTemp.RecordCount <> 0 Then
                str???? = zlCommFun.IncStr(str????)
            Else
                Exit Do
            End If
        Loop
    End If
    
    Dim strMembers As String
    strTemp = "": strMembers = ""
        
    If mbyt??ҩζ?? = 3 Then
        intGroup = 2
    ElseIf mbyt??ҩζ?? = 4 Then
        intGroup = 3
    End If
    With Me.msfRecipe
        For intCount = 1 To .Rows - 1
            For intFence = 0 To intGroup
                If Val(.TextMatrix(intCount, ?䷽?б?.???? * intFence + ?䷽?б?.ҩ??ID)) <> 0 Then   'id????Ϊ??
'                    If (Val(.TextMatrix(intCount, ?䷽?б?.???? * intFence + ?䷽?б?.????)) = 0) Then   '????????Ϊ0
'                        MsgBox "????????Ϊ0??????????????", vbInformation, gstrSysName
'                        .SetFocus
'                        .Row = intCount
'                        .Col = ?䷽?б?.???? * intFence + ?䷽?б?.????
'                        Exit Sub
'                    End If
                    
                    If strCheckID <> "" Then
                        If InStr(1, ";" & strCheckID & ";", ";" & Val(.TextMatrix(intCount, ?䷽?б?.???? * intFence + ?䷽?б?.ҩ??ID)) & "+" & Val(.TextMatrix(intCount, ?䷽?б?.???? * intFence + ?䷽?б?.????ID))) > 0 Then
                            MsgBox "?䷽?С?" & .TextMatrix(intCount, ?䷽?б?.???? * intFence + ?䷽?б?.????) & "???ظ?Ӧ?ã?", vbInformation, gstrSysName
                            .SetFocus
                            Exit Sub
                        End If
                    End If
                    
                    strCheckID = IIf(strCheckID = "", "", strCheckID & ";") & Val(.TextMatrix(intCount, ?䷽?б?.???? * intFence + ?䷽?б?.ҩ??ID)) & "+" & Val(.TextMatrix(intCount, ?䷽?б?.???? * intFence + ?䷽?б?.????ID))

                    strMembers = strMembers & "|" & Val(.TextMatrix(intCount, ?䷽?б?.???? * intFence + ?䷽?б?.ҩ??ID)) & _
                            "^" & IIf(Val(.TextMatrix(intCount, ?䷽?б?.???? * intFence + ?䷽?б?.????ID)) = 0, Null, Val(.TextMatrix(intCount, ?䷽?б?.???? * intFence + ?䷽?б?.????ID))) & _
                            "^" & Val(.TextMatrix(intCount, ?䷽?б?.???? * intFence + ?䷽?б?.????)) & _
                            "^" & Trim(.TextMatrix(intCount, ?䷽?б?.???? * intFence + ?䷽?б?.??ע))
                End If
            Next
        Next
        If strCheckID = "" Then MsgBox "δ?????䷽???ɣ?", vbInformation, gstrSysName: .SetFocus: Exit Sub
    End With
    strMembers = Mid(strMembers, 2)
    
    If cmbStationNo.Text = "" Then
        strվ?? = "Null"
    Else
        strվ?? = Mid(cmbStationNo.Text, 1, InStr(1, cmbStationNo.Text, "-") - 1)
    End If
    
    '???ݱ???
    If Me.Tag = "????" Then
        lngItemId = zlDatabase.GetNextId("??????ĿĿ¼")
'        If zlClinicCodeRepeat(Trim(Me.txt????.Text)) = True Then Exit Sub
    Else
        If zlClinicCodeRepeat(str????, lngItemId) = True Then Exit Sub
    End If
    gstrSql = lngItemId & "," & Me.txt????.Tag & ",'" & str???? & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txt????(0).Text) & "','" & Trim(Me.txtƴ??(0).Text) & "','" & Trim(Me.txt????(0).Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txt????(1).Text) & "','" & Trim(Me.txtƴ??(1).Text) & "','" & Trim(Me.txt????(1).Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txt˵??.Text) & "'"
    If Me.cboƵ??.Enabled Then
        gstrSql = gstrSql & ",'" & Left(Me.cboƵ??.Text, InStr(1, Me.cboƵ??.Text, "-") - 1) & "'"
    Else
        gstrSql = gstrSql & ",null"
    End If
    If Me.cbo?巨.Enabled Then
        gstrSql = gstrSql & "," & Me.cbo?巨.ItemData(Me.cbo?巨.ListIndex)
    Else
        gstrSql = gstrSql & ",0"
    End If
    If Me.cbo?÷?.Enabled Then
        gstrSql = gstrSql & "," & Me.cbo?÷?.ItemData(Me.cbo?÷?.ListIndex)
    Else
        gstrSql = gstrSql & ",0"
    End If
    gstrSql = gstrSql & "," & Val(Me.txt?Ƴ?.Text)
    
    gstrSql = gstrSql & "," & IIf(Me.txt?ο?.Tag = "", "Null", Me.txt?ο?.Tag)
    
    gstrSql = "zl_??ҩ?䷽_UPDATE(" & gstrSql & ",'" & strMembers & "'," & IIf(cmbStationNo.Visible = False Or cmbStationNo.Text = "", "Null", strվ??) & "," & intDosage & ")"
    
    Err = 0: On Error GoTo ErrHand
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
        
    '?????䷽???͵?ע????
    SaveSetting "ZLSOFT", "˽??ģ??\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "?䷽??̬", intDosage
    
    If Me.Tag = "????" Then
        If GetSetting("ZLSOFT", "????ģ??\" & App.ProductName & "\??????Ŀ????\", "????", 0) = 1 Then
            lngItemId = 0
            Call Form_Activate
            Me.txt????.SetFocus
            Exit Sub
        End If
    End If
    mblnOK = True
    Unload Me
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd?ο?_Click()
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = SelectRefer
    If Not rsTmp Is Nothing Then
        Me.txt?ο? = rsTmp("????"): Me.txt?ο?.Tag = rsTmp("ID"): strRefer = Me.txt?ο?
    Else
        MsgBox "û???ҵ??ɲο?????Ŀ??", vbInformation, Me.Caption
    End If
End Sub

Private Function SelectRefer(Optional ByVal strName As String = "") As ADODB.Recordset
    Dim strSql As String, strSQLItem As String
    Dim rsTmp As New ADODB.Recordset, iAttr As Integer
    
    On Error GoTo ErrHand
    strSql = "Select ???? From ???Ʒ???Ŀ¼ Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngClassId)
    
    If rsTmp.EOF Then
        iAttr = -1
    Else
        iAttr = rsTmp(0)
    End If
    If Len(strName) = 0 Then
        strSql = "Select level as ????,0 As ĩ??,ID,?ϼ?ID,????,????,'' As ˵?? From ???Ʋο????? a" & _
            " Where ????=" & iAttr & _
            " Start With a.?ϼ?id Is Null Connect By Prior a.id=a.?ϼ?id " & _
            " Union All" & _
            " Select 999 as ????,1,ID,????ID,????,????,˵?? From ???Ʋο?Ŀ¼ a Where ????=" & iAttr & " Order By ????,????"
    Else
        strSQLItem = " From ???Ʋο?Ŀ¼ A,???Ʋο????? B" & _
            " Where A.ID=B.?ο?Ŀ¼ID And A.????=" & iAttr & _
            " And (Upper(A.????) Like '" & UCase(strName) & "%'" & _
            " Or Upper(A.????) Like '" & mstrMatch & UCase(strName) & "%'" & _
            " Or Upper(B.????) Like '" & mstrMatch & UCase(strName) & "%'" & _
            " Or Upper(B.????) Like '" & mstrMatch & UCase(strName) & "%')"

        strSql = "Select Distinct 0 As ĩ??,ID,?ϼ?ID,????,????,'' As ˵?? From ???Ʋο????? a" & _
            " Where ????=" & iAttr & _
            " Start With ID In (Select ????ID " & strSQLItem & ") Connect By Prior a.?ϼ?id=a.id " & _
            " Union All" & _
            " Select Distinct 1,A.ID,A.????ID,A.????,A.????,A.˵?? " & strSQLItem & " Order By ????"
    End If
    Set SelectRefer = zlDatabase.ShowSelect(Me, strSql, 2, "?ο?", , , , , True)
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub cmd????_Click()
    With Me.tvwClass
        .Left = Me.txt????.Left
        .Top = Me.txt????.Top + Me.txt????.Height
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
End Sub

Private Sub Form_Activate()
'??ȡִ????Ŀ????Ϣ
    Dim lngCol As Long
    Dim intBe As Integer
    
    Err = 0: On Error GoTo ErrHand
    
    If mblnDosage = True Then Exit Sub

    gstrSql = "select A.????,A.????,A.?걾??λ as ˵??,A.????ʱ??,nvl(A.????ʱ??,to_date('3000-01-01','YYYY-MM-DD')) as ????ʱ??," & _
              " A.?ο?Ŀ¼id,B.???? As ?ο?????,A.վ?? " & _
              " from ??????ĿĿ¼ A,???Ʋο?Ŀ¼ B" & _
              " where A.?ο?Ŀ¼Id = B.id(+) And A.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)

    With rsTemp
        Me.txt????.MaxLength = .Fields("????").DefinedSize
        If .RecordCount > 0 Then
            Me.txt????.Text = !????: Me.txt????(0).Text = !????: Me.txt˵??.Text = IIf(IsNull(!˵??), "", !˵??)
            Me.txt?ο?.Text = NVL(!?ο?????)
            Me.txt?ο?.Tag = NVL(!?ο?Ŀ¼ID)
            SetStationNo IIf(IsNull(!վ??), "", !վ??)
            strRefer = Me.txt?ο?.Text
        End If
    End With

'    gstrSql = "Select 0 As ????, b.????, a.Id, a.????, a.????, a.???㵥λ, b.????????, b.ҽ??????" & vbNewLine & _
'        "From ??????ĿĿ¼ A, ??????Ŀ???? B" & vbNewLine & _
'        "Where a.Id = b.??????Ŀid And b.?շ?ϸĿid Is Null And b.????????id = [1]" & vbNewLine & _
'        "Union All" & vbNewLine & _
'        "Select 1 As ????, b.????, a.Id, a.????, a.???? || '(' || a.???? || ')' ????, a.???㵥λ, b.????????, b.ҽ??????" & vbNewLine & _
'        "From ?շ???ĿĿ¼ A, ??????Ŀ???? B" & vbNewLine & _
'        "Where a.Id = b.?շ?ϸĿid And b.??????Ŀid Is Null And b.????????id = [1]" & vbNewLine & _
'        "Order By ????"

    gstrSql = "Select b.????, b.??????ĿId As ҩ??id, b.?շ?ϸĿId As ????id, a.????, c.????, a.???㵥λ, b.????????, b.ҽ?????? " & vbNewLine & _
        "From ??????ĿĿ¼ A, ??????Ŀ???? B, ?շ???ĿĿ¼ C " & vbNewLine & _
        "Where a.Id = b.??????Ŀid And b.?շ?ϸĿid = c.Id(+) And b.????????id = [1] " & vbNewLine & _
        "Order By b.???? "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
    
    With rsTemp
        Me.msfRecipe.ClearBill
        Do While Not .EOF
            If Me.msfRecipe.Rows - 1 < ((.AbsolutePosition - 1) \ mbyt??ҩζ??) + 1 Then Me.msfRecipe.Rows = Me.msfRecipe.Rows + 1
            intFence = (.AbsolutePosition - 1) Mod mbyt??ҩζ??
            Me.msfRecipe.TextMatrix((.AbsolutePosition - 1) \ mbyt??ҩζ?? + 1, intFence * ?䷽?б?.???? + ?䷽?б?.ҩ??ID) = !ҩ??ID
            Me.msfRecipe.TextMatrix((.AbsolutePosition - 1) \ mbyt??ҩζ?? + 1, intFence * ?䷽?б?.???? + ?䷽?б?.????ID) = IIf(IsNull(!????ID), 0, !????ID)
            Me.msfRecipe.TextMatrix((.AbsolutePosition - 1) \ mbyt??ҩζ?? + 1, intFence * ?䷽?б?.???? + ?䷽?б?.????) = !???? & IIf(IsNull(!????), "", "(" & !???? & ")")
            Me.msfRecipe.TextMatrix((.AbsolutePosition - 1) \ mbyt??ҩζ?? + 1, intFence * ?䷽?б?.???? + ?䷽?б?.????) = FormatEx(IIf(IsNull(!????????), 0, !????????), 2)
            Me.msfRecipe.TextMatrix((.AbsolutePosition - 1) \ mbyt??ҩζ?? + 1, intFence * ?䷽?б?.???? + ?䷽?б?.??λ) = IIf(IsNull(!???㵥λ), "", !???㵥λ)
            Me.msfRecipe.TextMatrix((.AbsolutePosition - 1) \ mbyt??ҩζ?? + 1, intFence * ?䷽?б?.???? + ?䷽?б?.??ע) = IIf(IsNull(!ҽ??????), "", !ҽ??????)
            .MoveNext
        Loop
    End With

    gstrSql = "select R.?÷?ID,R.????,R.Ƶ??,R.?Ƴ?" & _
              " from ?????÷????? R" & _
              " where R.??ĿID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)

    With rsTemp
        Do While Not .EOF
            If !???? = 0 Then
                For intCount = 0 To Me.cbo?÷?.ListCount - 1
                    If Me.cbo?÷?.ItemData(intCount) = !?÷?ID Then Me.cbo?÷?.ListIndex = intCount: Exit For
                Next
            End If
            If !???? = 1 Then
                For intCount = 0 To Me.cbo?巨.ListCount - 1
                    If Me.cbo?巨.ItemData(intCount) = !?÷?ID Then Me.cbo?巨.ListIndex = intCount: Exit For
                Next
            End If
            For intCount = 0 To Me.cboƵ??.ListCount - 1
                If Left(Me.cboƵ??.List(intCount), InStr(1, Me.cboƵ??.List(intCount), "-") - 1) = IIf(IsNull(!Ƶ??), "", !Ƶ??) Then
                    Me.cboƵ??.ListIndex = intCount: Exit For
                End If
            Next
            Me.txt?Ƴ?.Text = IIf(IsNull(!?Ƴ?), 0, !?Ƴ?)
            .MoveNext
        Loop
    End With

    If Me.Tag = "????" Then
        lngItemId = 0

        If Val(zlDatabase.GetPara(61, glngSys)) = 0 Then    '??????Ŀ????????ģʽ
            gstrSql = "select nvl(max(????),'0000000') as ????" & _
                      " From ??????ĿĿ¼"
            '            If rsTemp.State = adStateOpen Then rsTemp.Close
            '            Call SQLTest(App.ProductName, Me.Caption, gstrSql)
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Activate")
            '            Call SQLTest
            Me.txt????.Text = Right(String(10, "0") & Val(rsTemp!????) + 1, Len(rsTemp!????))
        Else
            strTemp = Mid(Me.txt????.Text, 2, InStr(1, Me.txt????.Text, "]") - 2)
            gstrSql = "select nvl(max(????),'0000000') as ????" & _
                      " From ??????ĿĿ¼" & _
                      " Where  ???? like [1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, "8" & strTemp & "%")
            Err = 0: On Error Resume Next
            Me.txt????.Text = "8" & strTemp & Right(String(10, "0") & Val(rsTemp!????) + 1, Len(rsTemp!????) - 1 - Len(strTemp))
        End If

        Me.txt????(0).Text = ""
        Me.txt?ο? = "": Me.txt?ο?.Tag = "": strRefer = ""
    Else
        gstrSql = "select ????,????,????,???? from ??????Ŀ???? where ??????ĿID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
        With rsTemp
            Do While Not .EOF
                If !???? = 1 And !???? = 1 Then Me.txtƴ??(0).Text = !????
                If !???? = 1 And !???? = 2 Then Me.txt????(0).Text = !????
                If !???? = 9 Then Me.txt????(1).Text = !????
                If !???? = 9 And !???? = 1 Then Me.txtƴ??(1).Text = !????
                If !???? = 9 And !???? = 2 Then Me.txt????(1).Text = !????
                .MoveNext
            Loop
        End With
    End If

    If Me.Tag = "????" Then
        Me.cmdOK.Visible = False
        Me.cmdCancel.Caption = "?ر?(&C)"
        Me.txt????.Enabled = False: Me.cmd????.Enabled = False
        Me.txt????.Enabled = False
        Me.txt????(0).Enabled = False: Me.txtƴ??(0).Enabled = False: Me.txt????(0).Enabled = False
        Me.txt????(1).Enabled = False: Me.txtƴ??(1).Enabled = False: Me.txt????(1).Enabled = False
        Me.txt˵??.Enabled = False
        Me.cboƵ??.Enabled = False: Me.cbo?巨.Enabled = False: Me.cbo?÷?.Enabled = False
        Me.txt?Ƴ?.Enabled = False: Me.msfRecipe.Active = False
        Me.txt?ο?.Enabled = False
        Me.cmd?ο?.Enabled = False
    End If
    
    '??????ɫ
    For lngCol = 0 To msfRecipe.Cols - 1
        If lngCol Mod ?䷽?б?.???? = 0 Then
            For intBe = 0 To ?䷽?б?.???? - 1
                msfRecipe.SetColColor lngCol + intBe, &H8000000F
            Next
            lngCol = lngCol + ?䷽?б?.????
        End If
    Next
    mblnLoad = True
    
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub IniStationNo()
    Dim dblHeight As Double
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
    
'    lblStationNo.Visible = False
'    cmbStationNo.Visible = False
'
'    If gstrNodeNo <> "-" Then
    On Error GoTo ErrHand
    lblStationNo.Visible = True
    cmbStationNo.Visible = True
    
    strSql = "select ????,???? from zlnodelist"
    Set rsRecord = zlDatabase.OpenSQLRecord(strSql, "վ????ѯ")
    With cmbStationNo
        .AddItem ""
        Do While Not rsRecord.EOF
            .AddItem rsRecord!???? & "-" & rsRecord!????
            rsRecord.MoveNext
        Loop
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetStationNo(ByVal strNo As String)
    Dim n As Integer
    
'    If gstrNodeNo = "-" Then Exit Sub
    
    If strNo = "" Then
        cmbStationNo.ListIndex = 0
    Else
        For n = 1 To cmbStationNo.ListCount - 1
            If Mid(cmbStationNo.List(n), 1, InStr(1, cmbStationNo.List(n), "-") - 1) = strNo Then
                cmbStationNo.ListIndex = n
            End If
        Next
    End If
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    If Me.tvwClass.Visible Then
        Me.tvwClass.Visible = False: Me.txt????.SetFocus: Exit Sub
    ElseIf Me.lvwItems.Visible Then
        Me.lvwItems.Visible = False: Me.msfRecipe.SetFocus: Exit Sub
    End If
    Call cmdCancel_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call GetDefineSize
    Call IniStationNo
    
    ''??ҩ?䷽ÿ????ҩζ??
    mbyt??ҩζ?? = zlDatabase.GetPara(213, glngSys)
    With Me.msfRecipe
        strTemp = "   ?䷽???ɣ?  (*??˳??ѡ??ҩƷ?????뵥ζ????????Ҫʱ??д??????ע)"
        intCount = (.Width - Me.TextWidth(strTemp)) \ Me.TextWidth(Space(1))
        strTemp = strTemp & Space(intCount - 2)
        .Active = True
        .Rows = 2: .Cols = mbyt??ҩζ?? * ?䷽?б?.????
        .MsfObj.AllowUserResizing = flexResizeNone
        .MsfObj.ScrollBars = flexScrollBarBoth 'flexScrollBarVertical
        .MsfObj.GridColor = &H80000005: .MsfObj.BackColorBkg = &H80000005
        .MsfObj.MergeCells = flexMergeFree
        .MsfObj.MergeRow(0) = True
    
        .TxtCheck = True
        '        .TextMask = mstrChar
        For intCount = 0 To .Cols - 1
            .TextMatrix(0, intCount) = strTemp
            Select Case (intCount Mod ?䷽?б?.????)
            Case ?䷽?б?.?հ?
                .ColData(intCount) = 5: .ColWidth(intCount) = IIf(mbyt??ҩζ?? = 4, 200, 370)
            Case ?䷽?б?.ҩ??ID
                .ColData(intCount) = 5: .ColWidth(intCount) = 0
            Case ?䷽?б?.????ID
                .ColData(intCount) = 5: .ColWidth(intCount) = 0
            Case ?䷽?б?.????
                .ColData(intCount) = 1: .ColWidth(intCount) = IIf(mbyt??ҩζ?? = 4, 1500, 1700)
            Case ?䷽?б?.????
                .ColData(intCount) = 4: .ColWidth(intCount) = IIf(mbyt??ҩζ?? = 4, 500, 700)
            Case ?䷽?б?.??λ
                .ColData(intCount) = 5: .ColWidth(intCount) = IIf(mbyt??ҩζ?? = 4, 300, 400)
            Case ?䷽?б?.??ע
                .ColData(intCount) = 4: .ColWidth(intCount) = IIf(mbyt??ҩζ?? = 4, 850, 1000)
            End Select
        Next
'        .PrimaryCol = 3: .LocateCol = 3
        .Row = 1: .Col = 2
    End With

    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "????", "????", 1500
        .Add , "????", "????", 900
    End With
    With Me.lvwItems
        .Width = 2600
        .ColumnHeaders("????").Position = 1
        .SortKey = .ColumnHeaders("????").Index - 1
        .SortOrder = lvwAscending
    End With
    mstrMatch = IIf(GetSetting("ZLSOFT", "????ģ??\????", "????ƥ??", 0) = 0, "%", "")
    strRefer = ""
    
    mblnOK = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnDosage = False
    mblnLoad = False
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItems_DblClick()
    Dim lngCol As Long
    Dim intBe As Integer
    
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwItems
        If .SelectedItem.Icon = "ҩƷ" Then
            Me.msfRecipe.Text = .SelectedItem.Text
            Me.msfRecipe.TextMatrix(Me.msfRecipe.Row, Me.msfRecipe.Col - 1) = Mid(.SelectedItem.Key, 2)
            Me.msfRecipe.TextMatrix(Me.msfRecipe.Row, Me.msfRecipe.Col) = Me.msfRecipe.Text
            Me.msfRecipe.TextMatrix(Me.msfRecipe.Row, Me.msfRecipe.Col + 2) = .SelectedItem.Tag
        Else
            Me.msfRecipe.Text = .SelectedItem.Text
            Me.msfRecipe.TextMatrix(Me.msfRecipe.Row, Me.msfRecipe.Col) = Me.msfRecipe.Text
            
            With msfRecipe
                If Val(.TextMatrix(.Row, .Col - 4)) <> 0 Then 'idΪ?ղ?????????
                    If .Row = .Rows - 1 And .Col = .Cols - 1 Then
                        .Rows = .Rows + 1
                        .Row = .Rows - 1
                        .Col = 3
                    End If
                    
                    '??????ɫ
                    For lngCol = 0 To .Cols - 1
                        If lngCol Mod ?䷽?б?.???? = 0 Then
                            For intBe = 0 To ?䷽?б?.???? - 1
                                .SetColColor lngCol + intBe, &H8000000F
                            Next
                            lngCol = lngCol + ?䷽?б?.????
                        End If
                    Next
                    
                End If
            End With
        End If
        Me.msfRecipe.SetFocus
        Call zlCommFun.PressKey(vbKeyReturn)
    End With
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        Call lvwItems_DblClick
    End Select
End Sub

Private Sub lvwItems_LostFocus()
    Me.lvwItems.Visible = False
End Sub

Private Sub msfRecipe_CommandClick()
    Dim i As Integer
    Dim intDosageType As Integer
    Dim intCount As Integer
    Dim strName As String
    Dim intCurRow As Integer
    Dim intCurCol As Integer
    
'    If (Me.msfRecipe.Cols Mod 7) <> 2 Then Exit Sub
    For i = 0 To optDosageType.UBound
        If optDosageType(i).Value = True Then
            intDosageType = i
            Exit For
        End If
    Next
    
    intCurRow = msfRecipe.Row
    intCurCol = msfRecipe.Col
    mblnDosage = True
    frmMediDosage.ShowMe intDosageType, Me, "", strName
    
    If strName <> "" Then
        With msfRecipe
            If CheckDoubDosage(strName) = False Then
                MsgBox "??ҩƷ?Ѿ????ڣ??ظ?ҩƷ????¼?????Σ?", vbInformation, gstrSysName
                msfRecipe.Row = intCurRow
                msfRecipe.Col = intCurCol
                msfRecipe.SetFocus
                Exit Sub
            Else
                .Row = intCurRow
                .Col = intCurCol
                .SetFocus
                msfRecipe.TextMatrix(intCurRow, intCurCol - 2) = Split(strName, ",")(0) 'ҩ??id
                Me.msfRecipe.TextMatrix(intCurRow, intCurCol - 1) = Split(strName, ",")(1) '????id
                Me.msfRecipe.Text = Split(strName, ",")(2)  '????
                Me.msfRecipe.TextMatrix(intCurRow, intCurCol) = Me.msfRecipe.Text '????
                Me.msfRecipe.TextMatrix(intCurRow, intCurCol + 2) = Split(strName, ",")(3) '??λ
            End If
        End With
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CheckDoubDosage(ByVal strName As String) As Boolean
    '?????䷽??ҩƷ?Ƿ??ظ?
    Dim intRow As Integer
    Dim intFence As Integer
    Dim intGroup As Integer
    Dim strCheckID As String 'ҩ??ID+????ID
    
    If mbyt??ҩζ?? = 3 Then
        intGroup = 2
    ElseIf mbyt??ҩζ?? = 4 Then
        intGroup = 3
    End If
    
    strCheckID = Split(strName, ",")(0) & "+" & Split(strName, ",")(1)
    
    With msfRecipe
        For intRow = 1 To .Rows - 1
            For intFence = 0 To intGroup
                If Val(.TextMatrix(intRow, ?䷽?б?.???? * intFence + ?䷽?б?.ҩ??ID)) <> 0 Then  'id????Ϊ??
                    If strCheckID = .TextMatrix(intRow, ?䷽?б?.???? * intFence + ?䷽?б?.ҩ??ID) & "+" & .TextMatrix(intRow, ?䷽?б?.???? * intFence + ?䷽?б?.????ID) Then
                        Exit Function
                    End If
                    
'                    If InStr(1, Mid(strName, 1, InStr(3, strName, ",") - 1), .TextMatrix(intRow, ?䷽?б?.???? * intFence + 1) & "," & Trim(.TextMatrix(intRow, ?䷽?б?.???? * intFence + 2))) > 0 Then
'                        Exit Function
'                    End If
                End If
            Next
        Next
        CheckDoubDosage = True
    End With
End Function

Private Sub msfRecipe_EditChange(curText As String)
    mstrMedi = msfRecipe.TextMatrix(msfRecipe.Row, msfRecipe.Col)
End Sub

Private Sub msfRecipe_EditKeyDown(KeyCode As Integer, Shift As Integer)
'    If InStr("'`?|/\;,%", Chr(KeyCode)) > 0 Then KeyCode = 0
End Sub

Private Sub msfRecipe_EditKeyPress(KeyAscii As Integer)
    If InStr("'`?|/\;,%", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub msfRecipe_EnterCell(Row As Long, Col As Long)
    strInputed = Me.msfRecipe.TextMatrix(Row, Col)
End Sub

Private Sub msfRecipe_GotFocus()
    If Me.lvwItems.Visible Then Me.lvwItems.SetFocus
End Sub

Private Sub msfRecipe_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim i As Integer
    Dim intDosageType As Integer
    Dim intCount As Integer
    Dim strName As String
    Dim intCurRow As Integer
    Dim intCurCol As Integer
    Dim lngCol As Long
    Dim intBe As Integer
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.msfRecipe.Active = False Then Exit Sub
    With Me.msfRecipe
'        .TextMask = mstrChar
        Select Case (.Col Mod ?䷽?б?.????)
        Case ?䷽?б?.?հ?, ?䷽?б?.ҩ??ID, ?䷽?б?.????ID, ?䷽?б?.??λ
            Exit Sub
        Case ?䷽?б?.????
            If .TxtVisible = False Then
                If Trim(.TextMatrix(.Row, .Col)) = "" Then .TextMatrix(.Row, .Col) = "0"
                .TextMatrix(.Row, .Col) = FormatEx(.TextMatrix(.Row, .Col), 2)
            Else
                If Trim(.Text) = "" Then .Text = 0: .TextMatrix(.Row, .Col) = "0"
                .Text = FormatEx(.Text, 2)
            End If
'            If Val(.Text) = 0 Then
'                Cancel = True
'            End If
            Exit Sub
        Case ?䷽?б?.??ע
            .TextMask = ""
            If .TxtVisible = False Then
                If Trim(.TextMatrix(.Row, .Col)) = "" Then .TextMatrix(.Row, .Col) = Space(1)
                strTemp = UCase(Trim(.TextMatrix(.Row, .Col)))
            Else
                If Trim(.Text) = "" Then .Text = Space(1): .TextMatrix(.Row, .Col) = Space(1)
                strTemp = UCase(Trim(.Text))
            End If
            
            If strTemp = "" Or Not IsNumeric(strTemp) Then
                With msfRecipe
                    If Val(.TextMatrix(.Row, .Col - 4)) = 0 And Val(.TextMatrix(.Row, .Col - 5)) = 0 Then 'idΪ?ղ?????????
                    Else
                        If .Row = .Rows - 1 And .Col = .Cols - 1 Then
                            .Rows = .Rows + 1
                            .Row = .Rows - 1
                            .Col = 3
                            
                            '??????ɫ
                            For lngCol = 0 To .Cols - 1
                                If lngCol Mod ?䷽?б?.???? = 0 Then
                                    For intBe = 0 To ?䷽?б?.???? - 1
                                        .SetColColor lngCol + intBe, &H8000000F
                                    Next
                                    lngCol = lngCol + ?䷽?б?.????
                                End If
                            Next
                        End If
                    End If
                End With
                Exit Sub
            End If
            
            gstrSql = "select ????,????" & _
                    " from ??ҩ??????ע" & _
                    " where (???? like [1] or ???? like [2] or ???? like [2])" & _
                    " order by ????"
            Err = 0: On Error GoTo ErrHand
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strTemp & "%", gstrMatch & strTemp & "%")
            
            With rsTemp
                If .BOF Or .EOF Then Exit Sub
                If .RecordCount = 1 Then Me.msfRecipe.Text = !????: Me.msfRecipe.TextMatrix(Me.msfRecipe.Row, Me.msfRecipe.Col) = Me.msfRecipe.Text: Exit Sub
                Me.lvwItems.ListItems.Clear
                Do While Not .EOF
                    Set objItem = Me.lvwItems.ListItems.Add(, "_" & !????, !????)
                    objItem.Icon = "??ע": objItem.SmallIcon = "??ע"
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("????").Index - 1) = !????
                    .MoveNext
                Loop
                Me.lvwItems.ListItems(1).Selected = True
            End With
        Case ?䷽?б?.????
            If .TxtVisible = False Then
                If .TextMatrix(.Row, .Col) = "" Then
                    If .Col <> 3 Then Call zlCommFun.PressKey(vbKeyTab)
                    Exit Sub
                End If
                strTemp = UCase(Trim(.TextMatrix(.Row, .Col)))
            Else
                If Trim(.Text) = "" Then
                    If .Col <> 3 Then .SetFocus: Call zlCommFun.PressKey(vbKeyTab)
                    Exit Sub
                End If
                strTemp = UCase(Trim(.Text))
            End If
            If strInputed = strTemp Then Exit Sub
            
            For i = 0 To optDosageType.UBound
                If optDosageType(i).Value = True Then
                    intDosageType = i
                    Exit For
                End If
            Next
            
            intCurRow = msfRecipe.Row
            intCurCol = msfRecipe.Col
            mblnDosage = True
            frmMediDosage.ShowMe intDosageType, Me, strTemp, strName
            
            If strName <> "" Then
                With msfRecipe
                    If CheckDoubDosage(strName) = False Then
                        MsgBox "??ҩƷ?Ѿ????ڣ??ظ?ҩƷ????¼?????Σ?", vbInformation, gstrSysName
                        .Row = intCurRow
                        .Col = intCurCol
                        .SetFocus
                        Exit Sub
                    Else
                        .Row = intCurRow
                        .Col = intCurCol
                        .SetFocus
                        msfRecipe.TextMatrix(intCurRow, intCurCol - 2) = Split(strName, ",")(0) 'ҩ??id
                        Me.msfRecipe.TextMatrix(intCurRow, intCurCol - 1) = Split(strName, ",")(1) '????id
                        Me.msfRecipe.Text = Split(strName, ",")(2) '????
                        Me.msfRecipe.TextMatrix(intCurRow, intCurCol) = Me.msfRecipe.Text '????
                        Me.msfRecipe.TextMatrix(intCurRow, intCurCol + 2) = Split(strName, ",")(3) '??λ
                    End If
                End With
            Else
                .Text = mstrMedi
                Me.msfRecipe.TextMatrix(intCurRow, intCurCol) = .Text '????
            End If
            Exit Sub
        End Select
    End With
    
    With Me.lvwItems
        .Left = Me.msfRecipe.Left
        For intCount = 0 To Me.msfRecipe.Col - 1
            .Left = .Left + Me.msfRecipe.ColWidth(intCount)
        Next
        .Top = Me.msfRecipe.Top
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Cancel = True
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub msfRecipe_KeyPress(KeyAscii As Integer)
    If InStr("'`?|/\;,%", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub optDosageType_Click(Index As Integer)
    Dim intCol As Integer
    Dim intGroup As Integer
    Dim intFence As Integer
    Dim blnHaveData As Boolean
    
    If mblnLoad = True Then '???????????˺???????????
        If (mintOldShape = 1 And Index = 2) Or (mintOldShape = 2 And Index = 1) Then
        Else
            If mbyt??ҩζ?? = 3 Then
                intGroup = 2
            ElseIf mbyt??ҩζ?? = 4 Then
                intGroup = 3
            End If
            For intCount = 1 To msfRecipe.Rows - 1
                For intFence = 0 To intGroup
                    If Val(msfRecipe.TextMatrix(intCount, ?䷽?б?.???? * intFence + ?䷽?б?.????ID)) <> 0 Or Val(msfRecipe.TextMatrix(intCount, ?䷽?б?.???? * intFence + ?䷽?б?.ҩ??ID)) <> 0 Then   'id????Ϊ??
                        blnHaveData = True
                        Exit For
                    End If
                Next
                If blnHaveData = True Then
                    Exit For
                End If
            Next
            If blnHaveData = True And mblnClickNo = False Then
                If MsgBox("??̬?ı䣬???????䷽?????????ݣ??Ƿ???????", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    With msfRecipe
                        .Rows = 2
                        For intCol = 0 To .Cols - 1
                            .TextMatrix(1, intCol) = ""
                        Next
                    End With
                Else
                    mblnClickNo = True
                    optDosageType(mintOldShape).Value = True
                End If
            End If
        End If
    End If
End Sub

Private Sub optDosageType_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    
    mintOldShape = 0
    mblnClickNo = False
    For i = optDosageType.LBound To optDosageType.UBound
        If optDosageType(i).Value = True Then
            mintOldShape = i
            Exit For
        End If
    Next
End Sub

Private Sub tvwClass_DblClick()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Me.txt????.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
    Me.txt????.Text = Me.tvwClass.SelectedItem.Text
    Me.txt????.SetFocus
End Sub

Private Sub tvwClass_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
        If Me.tvwClass.SelectedItem.Children > 0 Then Exit Sub
        Call tvwClass_DblClick
    Case vbKeySpace
        Call tvwClass_DblClick
    Case vbKeyEscape
        Call tvwClass_LostFocus
    End Select
End Sub

Private Sub tvwClass_LostFocus()
    If Me.cmd???? Is ActiveControl Then Exit Sub
    Me.tvwClass.Visible = False
End Sub

Private Sub txt????_GotFocus()
    Me.txt????.SelStart = 0: Me.txt????.SelLength = 100
End Sub

Private Sub txt????_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt?ο?_GotFocus()
    Me.txt?ο?.SelStart = 0: Me.txt?ο?.SelLength = 100
End Sub


Private Sub txt?ο?_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        If Me.txt?ο? <> strRefer Then
            Set rsTmp = SelectRefer(Trim(Me.txt?ο?))
            If rsTmp Is Nothing Then
                Me.txt?ο? = strRefer
                Me.SetFocus
                MsgBox "û???ҵ??ɲο?????Ŀ??", vbInformation, Me.Caption
                Exit Sub
            Else
                Me.txt?ο? = rsTmp("????"): Me.txt?ο?.Tag = rsTmp("ID"): strRefer = Me.txt?ο?
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    If InStr(" ~!@#$%^&|=`;'""?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub


Private Sub txt?ο?_LostFocus()
    If Me.txt?ο? <> strRefer Then
        Me.txt?ο? = strRefer
    End If
End Sub


Private Sub txt????_GotFocus()
    Me.txt????.SelStart = 0: Me.txt????.SelLength = 100
End Sub

Private Sub txt????_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt?Ƴ?_GotFocus()
    Me.txt?Ƴ?.SelStart = 0: Me.txt?Ƴ?.SelLength = 100
End Sub

Private Sub txt?Ƴ?_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt????_GotFocus(Index As Integer)
    Me.txt????(Index).SelStart = 0: Me.txt????(Index).SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt????_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txt????(Index).Text = MoveSpecialChar(txt????(Index).Text)
        Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt????_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Me.txtƴ??(Index).Text = zlStr.GetCodeByORCL(Me.txt????(Index).Text, False, Me.txtƴ??(Index).MaxLength)
    Me.txt????(Index).Text = zlStr.GetCodeByORCL(Me.txt????(Index).Text, True, Me.txt????(Index).MaxLength)
End Sub

Private Sub txt????_LostFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtƴ??_GotFocus(Index As Integer)
    Me.txtƴ??(Index).SelStart = 0: Me.txtƴ??(Index).SelLength = 100
End Sub

Private Sub txtƴ??_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt˵??_GotFocus()
    Me.txt˵??.SelStart = 0: Me.txt˵??.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt˵??_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt˵??_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt????_GotFocus(Index As Integer)
    Me.txt????(Index).SelStart = 0: Me.txt????(Index).SelLength = 100
End Sub

Private Sub txt????_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub



