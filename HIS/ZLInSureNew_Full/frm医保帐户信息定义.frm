VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmҽ���ʻ���Ϣ���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�Զ��������Ϣ"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3930
   Icon            =   "frmҽ���ʻ���Ϣ����.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2700
      TabIndex        =   1
      Top             =   150
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   2700
      TabIndex        =   3
      Top             =   1290
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2700
      TabIndex        =   2
      Top             =   600
      Width           =   1100
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2850
      Top             =   3300
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
            Picture         =   "frmҽ���ʻ���Ϣ����.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LvwChoose 
      Height          =   3645
      Left            =   210
      TabIndex        =   0
      Top             =   630
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   6429
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   3175
      EndProperty
   End
   Begin VB.Label LblNote 
      Caption         =   "�ڴ˴�����ѡ���Լ�����Ȥ����Ϣ��"
      ForeColor       =   &H00800000&
      Height          =   450
      Left            =   240
      TabIndex        =   4
      Top             =   180
      Width           =   2040
   End
End
Attribute VB_Name = "frmҽ���ʻ���Ϣ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnOK As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim strColumns As String
    Dim intChoose As Integer
    
    With LvwChoose
        For intChoose = 1 To .ListItems.Count
            If .ListItems(intChoose).Checked Then
                strColumns = strColumns & IIf(strColumns = "", "", ",") & "'" & .ListItems(intChoose).Text & "'"
            End If
        Next
    End With
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\frmҽ���ʻ�", "�����ֶ�", strColumns
    mblnOK = True
    Unload Me
End Sub

Public Function SelectFields() As Boolean
    Dim strColumns As String
    Dim rsColumn As New ADODB.Recordset

    'ȡע���
    strColumns = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\frmҽ���ʻ�", "�����ֶ�", "")
    
    If strColumns = "" Then
        gstrSQL = " Select * From (" & _
                  " Select Distinct Column_Name ,1 State From All_Tab_Columns" & _
                  " Where Table_Name='������Ϣ'" & _
                  " And Not (Column_Name Like '%ID%')" & _
                  " And Column_Name Not In ('��������','������λ','����״��'))"
    Else
        gstrSQL = " Select * From (" & _
                  " Select Distinct Column_Name ,1 State From All_Tab_Columns" & _
                  " Where Table_Name='������Ϣ'" & _
                  " And Not (Column_Name Like '%ID%')" & _
                  " And Column_Name Not In ('��������','������λ','����״��')"
        gstrSQL = gstrSQL & " And Column_Name In (" & strColumns & ")"
        gstrSQL = gstrSQL & " Union " & _
                  " Select Distinct Column_Name ,0 State From All_Tab_Columns" & _
                  " Where Table_Name='������Ϣ'" & _
                  " And Not (Column_Name Like '%ID%')" & _
                  " And Column_Name Not In ('��������','������λ','����״��'" & IIf(strColumns = "", "", "," & strColumns) & "))" & _
                  " Order by State desc,Column_Name"
    End If
    Call OpenRecordset(rsColumn, Me.Caption)
    
    If rsColumn.EOF Then
        MsgBox "���ݱ�����������ʹ�øù��ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    'װ��������
    LvwChoose.ListItems.Clear
    strColumns = "," & strColumns & ","
    
    With rsColumn
        Do While Not .EOF
            LvwChoose.ListItems.Add , "K_" & LvwChoose.ListItems.Count + 1, !Column_Name, 1, 1
            If InStr(1, strColumns, ",'" & !Column_Name & "',") <> 0 Then LvwChoose.ListItems(LvwChoose.ListItems.Count).Checked = True
            .MoveNext
        Loop
    End With
        
    mblnOK = False
    frmҽ���ʻ���Ϣ����.Show vbModal, frmҽ���ʻ�
    SelectFields = mblnOK
End Function
