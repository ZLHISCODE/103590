VERSION 5.00
Begin VB.Form frmMain_北京尚洋对码上传 
   Caption         =   "对码上传"
   ClientHeight    =   2070
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4935
   Icon            =   "frmMain_北京尚洋对码上传.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4935
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "关闭(&C)"
      Height          =   350
      Left            =   3610
      TabIndex        =   1
      Top             =   1500
      Width           =   1100
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "上传(&O)"
      Height          =   350
      Left            =   2350
      TabIndex        =   0
      Top             =   1485
      Width           =   1100
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "终止(&S)"
      Height          =   350
      Left            =   2350
      TabIndex        =   2
      Top             =   1500
      Width           =   1100
   End
   Begin VB.Label LabCaption 
      AutoSize        =   -1  'True
      Caption         =   "上传医保对码信息"
      Height          =   180
      Left            =   255
      TabIndex        =   6
      Top             =   135
      Width           =   1440
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      X1              =   195
      X2              =   4710
      Y1              =   1095
      Y2              =   1095
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   195
      X2              =   4740
      Y1              =   645
      Y2              =   645
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      X1              =   195
      X2              =   4710
      Y1              =   1110
      Y2              =   1110
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      X1              =   195
      X2              =   4740
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Label LabStatus 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12.25%"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   195
      TabIndex        =   4
      Top             =   780
      Width           =   4560
   End
   Begin VB.Label pbrBar 
      Height          =   240
      Left            =   165
      TabIndex        =   3
      Top             =   2775
      Visible         =   0   'False
      Width           =   4560
   End
   Begin VB.Label labBar 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   195
      TabIndex        =   5
      Top             =   750
      Width           =   4560
   End
End
Attribute VB_Name = "frmMain_北京尚洋对码上传"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mintInsure              As Integer
Private mblnStop                As Boolean
Dim lngLoop                     As Integer
Dim strSQL                      As String

Const strUpSql1 = "Select" & vbNewLine & _
                "B.类别 As ITEM_CLASS," & vbNewLine & _
                "B.编码 As ITEM_CODE," & vbNewLine & _
                "B.名称 AS ITEM_NAME," & vbNewLine & _
                "SubstrB(B.规格,1,50) AS ITEM_SPECIFICATION," & vbNewLine & _
                "SubstrB(B.计算单位,1,8) AS UNIT," & vbNewLine & _
                "nvl(C.单价,0) AS STANDARD_PRICE," & vbNewLine & _
                "zl_split(zl_split(A.附注,'|||',1),'.',0) AS ITEM_ON_DISPENSARY_RECEIPT," & vbNewLine & _
                "zl_split(zl_split(A.附注,'|||',2),'.',0) AS ITEM_ON_RESIDENT_RECEIPT," & vbNewLine & _
                "Null AS ITEM_NO_DEPT_STAT," & vbNewLine & _
                "Null AS ITEM_NO_ACCOUNTANT_ITEM," & vbNewLine & _
                "Null AS MEMO," & vbNewLine & _
                "C.执行日期 AS START_DATE," & vbNewLine & _
                "C.终止日期 AS STOP_DATE," & vbNewLine & _
                "C.调价人 AS OPERATOR," & vbNewLine & _
                "B.建档时间 AS MODIFY_DATE," & vbNewLine & _
                "A.项目编码 AS COLLATE_RELATION," & vbNewLine & _
                "1 AS CONVERSION_RATE," & vbNewLine & _
                "Null AS ITEM_FORM," & vbNewLine & _
                "1 AS CHRONIC_CONVERSION_RATE," & vbNewLine
Const strUpSql2 = "Null AS CHRONIC_MIN_UNIT," & vbNewLine & _
                "Null AS EXAMINE_PERSON," & vbNewLine & _
                "Null AS EXAMINE_DATE," & vbNewLine & _
                "'0' AS EXAMINE_FLAG," & vbNewLine & _
                "DECODE(B.类别,'5','01','6','02','7','03','00') AS gkfldm," & vbNewLine & _
                "'0000' AS kzyfdm," & vbNewLine & _
                "NULL AS zxks," & vbNewLine & _
                "DECODE(B.类别,'5','0','6','0','7','0','1') AS ypjcbz," & vbNewLine & _
                "D.简码 AS pydm," & vbNewLine & _
                "NULL AS zxksmc" & vbNewLine & _
                "from 保险支付项目 A,收费细目 B,(Select x.收费细目id, y.现价 As 单价,y.执行日期,y.终止日期,y.调价人 From (Select 收费细目id, Max(ID) As ID  From 收费价目   Where Sysdate >= 执行日期 And Sysdate <= 终止日期  Group By 收费细目id) X, 收费价目 Y  Where x.Id = y.Id) C,收费项目别名 D" & vbNewLine & _
                "Where A.收费细目ID = B.ID And A.收费细目ID = C.收费细目id And A.收费细目ID=D.收费细目ID And A.险类=92 And D.性质=1 And D.码类=1" & vbNewLine
Const strUpSql = strUpSql1 & strUpSql2

Public Property Let intinsure(ByVal vNewValue As Integer)
    mintInsure = vNewValue
End Property

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdUp_Click()
    Dim strErrMsg           As String   '消息
    Dim lngCount            As Long
    Dim strRem              As String
    Dim strCodeID           As String
    Dim strCodeName         As String
    Dim rsTemp              As ADODB.Recordset
    Dim rsCenter            As ADODB.Recordset
    Dim lngID               As Long
    Dim strOutArray()       As String
    Dim lngOut              As Long
On Error GoTo ErrH
    lngOut = 0
    mblnStop = False
    cmdUp.Visible = False
    cmdCancel.Enabled = False
    cmdStop.Visible = True
    cmdStop.Enabled = False
    labSTATUS.Caption = "正在读取数据..."
    DoEvents
    gstrSQL = strUpSql
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If Not (rsTemp.EOF Or rsTemp.BOF) Then
        ReDim strOutArray(rsTemp.RecordCount - 1) As String
        labBar.Visible = True
        Do While Not (rsTemp.EOF Or rsTemp.BOF)
            labSTATUS.Caption = Format(Round(rsTemp.Bookmark / rsTemp.RecordCount * 100, 2), "0.00") & " %"
            labBar.Width = rsTemp.Bookmark * pbrBar.Width / rsTemp.RecordCount
            DoEvents
            '检测当前编码是否在中心存在
            gstrSQL = "Select EXAMINE_FLAG From PRICELIST_DICT Where ITEM_CODE='" & rsTemp!ITEM_CODE & "'"
            Set rsCenter = gcn尚洋.Execute(gstrSQL)
            If rsCenter.EOF Or rsCenter.BOF Then
                '中心不存在数据直接新增
                gstrSQL = "INSERT INTO PRICELIST_DICT([ITEM_CLASS],[ITEM_CODE],[ITEM_NAME],[ITEM_SPECIFICATION],[UNIT],[STANDARD_PRICE],[ITEM_ON_DISPENSARY_RECEIPT],[ITEM_ON_RESIDENT_RECEIPT],[ITEM_NO_DEPT_STAT],[ITEM_NO_ACCOUNTANT_ITEM],[MEMO],[START_DATE],[STOP_DATE],[OPERATOR],[MODIFY_DATE],[COLLATE_RELATION],[CONVERSION_RATE],[ITEM_FORM],[CHRONIC_CONVERSION_RATE],[CHRONIC_MIN_UNIT],[EXAMINE_PERSON],[EXAMINE_DATE],[EXAMINE_FLAG],[gkfldm],[kzyfdm],[zxks],[ypjcbz],[pydm],[zxksmc]) Values(" & vbNewLine & _
                          "'" & rsTemp!ITEM_CLASS & "','" & rsTemp!ITEM_CODE & "','" & rsTemp!ITEM_NAME & "','" & rsTemp!ITEM_SPECIFICATION & "','" & rsTemp!UNIT & "','" & rsTemp!STANDARD_PRICE & "','" & rsTemp!ITEM_ON_DISPENSARY_RECEIPT & "','" & rsTemp!ITEM_ON_RESIDENT_RECEIPT & "','" & rsTemp!ITEM_NO_DEPT_STAT & "','" & rsTemp!ITEM_NO_ACCOUNTANT_ITEM & "','" & rsTemp!Memo & "','" & rsTemp!START_DATE & "','" & rsTemp!STOP_DATE & "','" & rsTemp!OPERATOR & "','" & rsTemp!MODIFY_DATE & "','" & rsTemp!COLLATE_RELATION & "','" & rsTemp!CONVERSION_RATE & "','" & rsTemp!ITEM_FORM & "','" & rsTemp!CHRONIC_CONVERSION_RATE & "','" & rsTemp!CHRONIC_MIN_UNIT & "',NULL,NULL,'" & rsTemp!EXAMINE_FLAG & "','" & rsTemp!gkfldm & "','" & rsTemp!kzyfdm & "','" & rsTemp!zxks & "','" & rsTemp!ypjcbz & "','" & rsTemp!pydm & "','" & rsTemp!zxksmc & "')"
                Call gcn尚洋.Execute(gstrSQL)
            ElseIf rsCenter.RecordCount = 0 Then
                '新增数据
                gstrSQL = "INSERT INTO PRICELIST_DICT([ITEM_CLASS],[ITEM_CODE],[ITEM_NAME],[ITEM_SPECIFICATION],[UNIT],[STANDARD_PRICE],[ITEM_ON_DISPENSARY_RECEIPT],[ITEM_ON_RESIDENT_RECEIPT],[ITEM_NO_DEPT_STAT],[ITEM_NO_ACCOUNTANT_ITEM],[MEMO],[START_DATE],[STOP_DATE],[OPERATOR],[MODIFY_DATE],[COLLATE_RELATION],[CONVERSION_RATE],[ITEM_FORM],[CHRONIC_CONVERSION_RATE],[CHRONIC_MIN_UNIT],[EXAMINE_PERSON],[EXAMINE_DATE],[EXAMINE_FLAG],[gkfldm],[kzyfdm],[zxks],[ypjcbz],[pydm],[zxksmc]) Values(" & vbNewLine & _
                          "'" & rsTemp!ITEM_CLASS & "','" & rsTemp!ITEM_CODE & "','" & rsTemp!ITEM_NAME & "','" & rsTemp!ITEM_SPECIFICATION & "','" & rsTemp!UNIT & "','" & rsTemp!STANDARD_PRICE & "','" & rsTemp!ITEM_ON_DISPENSARY_RECEIPT & "','" & rsTemp!ITEM_ON_RESIDENT_RECEIPT & "','" & rsTemp!ITEM_NO_DEPT_STAT & "','" & rsTemp!ITEM_NO_ACCOUNTANT_ITEM & "','" & rsTemp!Memo & "','" & rsTemp!START_DATE & "','" & rsTemp!STOP_DATE & "','" & rsTemp!OPERATOR & "','" & rsTemp!MODIFY_DATE & "','" & rsTemp!COLLATE_RELATION & "','" & rsTemp!CONVERSION_RATE & "','" & rsTemp!ITEM_FORM & "','" & rsTemp!CHRONIC_CONVERSION_RATE & "','" & rsTemp!CHRONIC_MIN_UNIT & "',NULL,NULL,'" & rsTemp!EXAMINE_FLAG & "','" & rsTemp!gkfldm & "','" & rsTemp!kzyfdm & "','" & rsTemp!zxks & "','" & rsTemp!ypjcbz & "','" & rsTemp!pydm & "','" & rsTemp!zxksmc & "')"
                Call gcn尚洋.Execute(gstrSQL)
            ElseIf rsCenter!EXAMINE_FLAG = 1 Then
                '中心已审核不能修改
                strOutArray(lngOut) = "【" & rsTemp!ITEM_CODE & "】中心已审核，上传失败！"
                lngOut = lngOut + 1
            ElseIf rsCenter!EXAMINE_FLAG = 2 Then
                '中心已上传不能修改
                strOutArray(lngOut) = "【" & rsTemp!ITEM_CODE & "】中心已上传，上传失败！"
                lngOut = lngOut + 1
            Else
                '中心存在数据【先删除中心数据】
                gstrSQL = "Delete PRICELIST_DICT Where ITEM_CODE='" & rsTemp!ITEM_CODE & "'"
                Call gcn尚洋.Execute(gstrSQL)
                '新增数据
                gstrSQL = "INSERT INTO PRICELIST_DICT([ITEM_CLASS],[ITEM_CODE],[ITEM_NAME],[ITEM_SPECIFICATION],[UNIT],[STANDARD_PRICE],[ITEM_ON_DISPENSARY_RECEIPT],[ITEM_ON_RESIDENT_RECEIPT],[ITEM_NO_DEPT_STAT],[ITEM_NO_ACCOUNTANT_ITEM],[MEMO],[START_DATE],[STOP_DATE],[OPERATOR],[MODIFY_DATE],[COLLATE_RELATION],[CONVERSION_RATE],[ITEM_FORM],[CHRONIC_CONVERSION_RATE],[CHRONIC_MIN_UNIT],[EXAMINE_PERSON],[EXAMINE_DATE],[EXAMINE_FLAG],[gkfldm],[kzyfdm],[zxks],[ypjcbz],[pydm],[zxksmc]) Values(" & vbNewLine & _
                          "'" & rsTemp!ITEM_CLASS & "','" & rsTemp!ITEM_CODE & "','" & rsTemp!ITEM_NAME & "','" & rsTemp!ITEM_SPECIFICATION & "','" & rsTemp!UNIT & "','" & rsTemp!STANDARD_PRICE & "','" & rsTemp!ITEM_ON_DISPENSARY_RECEIPT & "','" & rsTemp!ITEM_ON_RESIDENT_RECEIPT & "','" & rsTemp!ITEM_NO_DEPT_STAT & "','" & rsTemp!ITEM_NO_ACCOUNTANT_ITEM & "','" & rsTemp!Memo & "','" & rsTemp!START_DATE & "','" & rsTemp!STOP_DATE & "','" & rsTemp!OPERATOR & "','" & rsTemp!MODIFY_DATE & "','" & rsTemp!COLLATE_RELATION & "','" & rsTemp!CONVERSION_RATE & "','" & rsTemp!ITEM_FORM & "','" & rsTemp!CHRONIC_CONVERSION_RATE & "','" & rsTemp!CHRONIC_MIN_UNIT & "',NULL,NULL,'" & rsTemp!EXAMINE_FLAG & "','" & rsTemp!gkfldm & "','" & rsTemp!kzyfdm & "','" & rsTemp!zxks & "','" & rsTemp!ypjcbz & "','" & rsTemp!pydm & "','" & rsTemp!zxksmc & "')"
                Call gcn尚洋.Execute(gstrSQL)
            End If
            
            rsTemp.MoveNext
        Loop
    Else
        labBar.Visible = False
        labSTATUS.Caption = "没有对码数据不能上传！"
    End If
    '写入日志
    writeOut "c:\Up" & UserInfo.姓名 & ".txt", strOutArray
    cmdStop.Visible = True
    cmdUp.Visible = True
    cmdCancel.Enabled = True
    Exit Sub
ErrH:
    cmdUp.Visible = True
    cmdCancel.Enabled = True
End Sub

Private Sub cmdStop_Click()
    mblnStop = True
End Sub

Private Sub Form_Load()
    labBar.Width = 0
    labSTATUS.Caption = ""
End Sub








