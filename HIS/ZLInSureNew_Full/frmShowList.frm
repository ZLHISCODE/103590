VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmShowList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ѡ����"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmShowList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2070
      Top             =   1320
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
            Picture         =   "frmShowList.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LvwSelect 
      Height          =   3195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   5636
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "�Һ���ˮ��"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "�Һſ���"
         Object.Width           =   3528
      EndProperty
   End
End
Attribute VB_Name = "frmShowList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private arrRow, arrCol, arrTitle
Private strData As String, strTitle
Private intRow As Integer, intCol As Integer
Private Const strSplit As String = "||"
'���е�������||�ָ������е�������;�ָ�

Private Sub Form_Load()
    '������û������
    Dim lvwItem As ListItem
    arrCol = Split(strData, strSplit)
    arrTitle = Split(strTitle, strSplit)
    LvwSelect.ListItems.Clear
    
    For intCol = 0 To UBound(arrCol)
        arrRow = Split(arrCol(intCol), ";")
        For intRow = 0 To UBound(arrRow)
            If intCol = 0 Then
                LvwSelect.ListItems.Add , "K" & LvwSelect.ListItems.Count + 1, arrRow(intRow), , 1
            Else
                LvwSelect.ListItems("K" & intRow + 1).SubItems(intCol) = arrRow(intRow)
            End If
        Next
        LvwSelect.ColumnHeaders(intCol + 1).Text = arrTitle(intCol)
    Next
End Sub

Public Function ShowME(ByVal IN_strData As String, Optional ByVal IN_strTitle As String = "�Һ���ˮ��||�Һſ���") As String
    strData = IN_strData
    strTitle = IN_strTitle
    Me.Show 1
    ShowME = strData
End Function

Private Sub LvwSelect_DblClick()
    Call LvwSelect_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub LvwSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        '�����������ݣ�������û�����ݣ�
        strData = LvwSelect.SelectedItem.Text
        For intCol = 1 To LvwSelect.ColumnHeaders.Count - 1
            strData = strData & IIf(strData = "", "", ";") & LvwSelect.SelectedItem.SubItems(intCol)
        Next
        Unload Me
        Exit Sub
    End If
End Sub

