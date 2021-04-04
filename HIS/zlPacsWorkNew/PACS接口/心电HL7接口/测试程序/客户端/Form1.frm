VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   7200
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtInfo 
      Height          =   4815
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   120
      Width           =   5415
   End
   Begin VB.TextBox txtIP 
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Text            =   "127.0.0.1"
      Top             =   5160
      Width           =   1575
   End
   Begin VB.TextBox txtPort 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Text            =   "104"
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "发送"
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "断开"
      Height          =   495
      Left            =   5640
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "请求"
      Height          =   495
      Left            =   5640
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6600
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   ""
   End
   Begin VB.Label lblLocalPort 
      Height          =   495
      Left            =   3360
      TabIndex        =   8
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "IP地址："
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "端口："
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   5640
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    
    If Me.Winsock1.State = sckClosed Then
        Me.Winsock1.RemoteHost = Me.txtIP.Text
        Me.Winsock1.RemotePort = Val(Me.txtPort.Text)
        Me.Winsock1.LocalPort = 0   '客户端的本地端口指定为0，在连接的时候，会自动使用空闲窗口，避免因为断开之后，无法再次连接
        Me.Winsock1.Connect
    Else
        Me.Winsock1.Close
        Me.Winsock1.RemoteHost = Me.txtIP.Text
        Me.Winsock1.RemotePort = Val(Me.txtPort.Text)
        Me.Winsock1.LocalPort = 0
        Me.Winsock1.Connect
    End If
    If Me.Winsock1.State = sckConnecting Then
        Me.txtInfo.Text = Me.txtInfo.Text & "等待连接，IP：" & Me.txtIP.Text & " PORT:" & Me.txtPort.Text & vbCrLf
        lblLocalPort = Me.Winsock1.LocalPort
    End If
End Sub

Private Sub Command2_Click()
    Me.Winsock1.Close
    If Me.Winsock1.State = sckClosed Then
        Me.txtInfo.Text = Me.txtInfo.Text & "断开连接，IP：" & Me.txtIP.Text & " PORT:" & Me.txtPort.Text & vbCrLf
    End If
End Sub

Private Sub Command3_Click()
    Dim strData As String
    Dim bTemp As Byte
    bTemp = 10
    
    '连接GE的数据，chr(11)开头，chr(13)连接；chr(28)结束
    
    
'    strData = Chr$(11) & "MSH|^~\&|MESA_ADT|XYZ_ADMITTING|MESA_IS|XYZ_HOSPITAL|||ADT^A04|101102|P|2.4|" & vbCrLf _
'                            & "EVN||200004211000||||200004210950 " & vbCrLf _
'                            & "PID|||583070^^^ADT1||BLACK^CHARLES||19980704|M||AI|7616 STANFORD AVE^^ST. LOUIS^MO^63130|||||||20-98-1701||||||||||||" & vbCrLf _
'                            & "PV1||E||||||5101^NELL^FREDERICK^P^^DR|||||||||||V1002^^^ADT1|||||||||||||||||||||||||200004210950||||||||" & vbCrLf _
'                            & Chr(28) & Chr(13)
                            
'    strData = Chr$(11) & "MSH|^~\&|MESA_ADT|XYZ_ADMITTING|MESA_IS|XYZ_HOSPITAL|||ADT^A04|101102|P|2.4|" & vbCrLf _
'                            & "EVN||200004211000||||200004210950 " & vbCrLf _
'                            & "PID|||583070^^^ADT1||BLACK^CHARLES||19980704|M||AI|7616 STANFORD AVE^^ST. LOUIS^MO^63130|||||||20-98-1701||||||||||||" & vbCrLf _
'                            & "PV1||E||||||5101^NELL^FREDERICK^P^^DR|||||||||||V1002^^^ADT1|||||||||||||||||||||||||200004210950||||||||" & vbCrLf _
'                            & Chr(28) & Chr(13)
                            
'    strData = Chr$(11) & "MSH|^~\&|MESA_ADT|XYZ_ADMITTING|MESA_IS|XYZ_HOSPITAL|||ADT^A04|101102|P|2.3.1||||||||" & vbCrLf _
'                        & "EVN||200004211000||||200004210950 " & vbCrLf _
'                        & "PID|||583020^^^ADT1||WHITE^CHARLES||19980704|M||AI|7616 STANFORD AVE^^ST. LOUIS^MO^63130|||||||20-98-1701||||||||||||" & vbCrLf _
'                        & "PV1||E||||||5101^NELL^FREDERICK^P^^DR|||||||||||V1002^^^ADT1|||||||||||||||||||||||||200004210950||||||||" & vbCrLf _
'                        & Chr(28) & Chr(13)
     'Me.Winsock1.SendData Chr$(11)
'

'普通ORU_RO1消息
'strData = Chr$(11) & "MSH|^~\&|MUSE ECG Result 1|MEI MUSE|CCG-MUSE Results|CCG|20120222115059||ORU^R01|20120222115059|P|2.4" & Chr(13) & _
'                "PID|1|0000001202220054|||柴^宏勋||19720220|M||U|^^^^^||||||||" & Chr(13) & _
'                "PV1|1||(65535)^^^SITE0001^^^^^||||^^^^^^^^^^^|^^^^^^^^^^^|^^^^^^^^^^^||||||||^^^^^^^^^^^|||||||||||||||||||||||||||||||||" & Chr(13) & _
'                "OBR|1|123091||^^^^12 Lead ECG||20120222115059|20120222101743^10%2017%2043%200||||||^|20120222101126||^^^^^^^^^^^||||1661||20120222115059|||F|||||||^郭^勇娟^^^^^^^^1^|^^^^^^^^^^^|^^^^^^^^^^^|^郭^勇娟^^^^^^^^1^|" & Chr(13) & _
'                "DG1|1|||" & Chr(13) & "DG1|2||" & Chr(13) & "DG1|3||" & Chr(13) & "NTE|1||" & Chr(13) & "ZEX|1|^^^|^^^|^^^|" & Chr(13) & _
'                "ZPH|1||||||||||^|^|^|^|^" & Chr(13) & "OBX|1|ST|903^Acquisition Device||MAC80||||||F" & Chr(13) & "OBX|2|ST|550^Systolic BP|||mmHg|||||F" & Chr(13) & "OBX|3|ST|551^Diastolic BP|||mmHg|||||F" & Chr(13) & _
'                "OBX|4|ST|552^Ventricular Rate||79|BPM|||||F" & Chr(13) & "OBX|5|ST|553^Atrial Rate||79|BPM|||||F" & Chr(13) & "OBX|6|ST|554^P-R Interval||150|ms|||||F" & Chr(13) & "OBX|7|ST|555^QRS Duration||86|ms|||||F" & Chr(13) & _
'                "OBX|8|ST|556^Q-T Interval||394|ms|||||F" & Chr(13) & "OBX|9|ST|557^QTC Calculation(Bezet)||451|ms|||||F" & Chr(13) & "OBX|10|ST|558^P Axis||65|degrees|||||F" & Chr(13) & "OBX|11|ST|568^Calculated P Axis||65|degrees|||||F" & Chr(13) & _
'                "OBX|12|ST|559^R Axis||13|degrees|||||F" & Chr(13) & "OBX|13|ST|569^Calculated R Axis||13|degrees|||||F" & Chr(13) & "OBX|14|ST|560^T Axis||34|degrees|||||F" & Chr(13) & "OBX|15|ST|570^Calculated T Axis||34|degrees|||||F" & Chr(13) & _
'                "OBX|16|ST|561^QRS Count||13|beats|||||F" & Chr(13) & "OBX|17|ST|562^Q Onset||226|ms|||||F" & Chr(13) & "OBX|18|ST|563^Q Offset||269|ms|||||F" & Chr(13) & "OBX|19|ST|564^P Onset||151|ms|||||F" & Chr(13) & _
'                "OBX|20|ST|565^P Offset||202|ms|||||F" & Chr(13) & "OBX|21|ST|566^T Offset||423|ms|||||F" & Chr(13) & "OBX|22|ST|575^QTC Fredericia||432|ms|||||F" & Chr(13) & "OBX|23|ST|576^QTC Framingham||432|ms|||||F" & Chr(13) & _
'                "OBX|24|ST|578^QTC RR|||ms|||||F" & Chr(13) & "OBX|25|ST|15006^MAC1200 PP Interval||759|ms|||||F" & Chr(13) & "OBX|26|ST|15002^MAC1200 RR Interval||758|ms|||||F" & Chr(13) & "OBX|27|RP|MUSEWebURL||http://ELECTROCARDDIOG/musescripts/museweb.dll?RetrieveTestByDateTime?PatientID=0000001202220054\T\Date=22-02-2012\T\Time=10%3a17%3a43%3a00\T\TestType=ECG\T\Site=1\T\OutputType=PDF\T\Ext=PDF||||||F" & Chr(13) & _
'                Chr(28) & Chr(13)

'超长ORU_R01消息
strData = Chr$(11) & "MSH|^~\&|MUSE ECG Result 1|MEI MUSE|CCG-MUSE Results|CCG|20120222115059||ORU^R01|20120222115059|P|2.4" & Chr(13) & _
                "PID|1|0000001202220054|||柴^宏勋||19720220|M||U|^^^^^||||||||" & Chr(13) & "PV1|1||(65535)^^^SITE0001^^^^^||||^^^^^^^^^^^|^^^^^^^^^^^|^^^^^^^^^^^||||||||^^^^^^^^^^^|||||||||||||||||||||||||||||||||" & Chr(13) & _
                "OBR|1|123091||^^^^12 Lead ECG||20120222115059|20120222101743^10%2017%2043%200||||||^|20120222101126||^^^^^^^^^^^||||1661||20120222115059|||F|||||||^郭^勇娟^^^^^^^^1^|^^^^^^^^^^^|^^^^^^^^^^^|^郭^勇娟^^^^^^^^1^|" & Chr(13) & _
                "DG1|1|||" & Chr(13) & "DG1|2||" & Chr(13) & "DG1|3||" & Chr(13) & "NTE|1||" & Chr(13) & "ZEX|1|^^^|^^^|^^^|" & Chr(13) & _
                "ZPH|1||||||||||^|^|^|^|^" & Chr(13) & "OBX|1|ST|903^Acquisition Device||MAC80||||||F" & Chr(13) & "OBX|2|ST|550^Systolic BP|||mmHg|||||F" & Chr(13) & "OBX|3|ST|551^Diastolic BP|||mmHg|||||F" & Chr(13) & _
                "OBX|4|ST|552^Ventricular Rate||79|BPM|||||F" & Chr(13) & "OBX|5|ST|553^Atrial Rate||79|BPM|||||F" & Chr(13) & "OBX|6|ST|554^P-R Interval||150|ms|||||F" & Chr(13) & "OBX|7|ST|555^QRS Duration||86|ms|||||F" & Chr(13) & _
                "OBX|8|ST|556^Q-T Interval||394|ms|||||F" & Chr(13) & "OBX|9|ST|557^QTC Calculation(Bezet)||451|ms|||||F" & Chr(13) & "OBX|10|ST|558^P Axis||65|degrees|||||F" & Chr(13) & "OBX|11|ST|568^Calculated P Axis||65|degrees|||||F" & Chr(13) & _
                "OBX|12|ST|559^R Axis||13|degrees|||||F" & Chr(13) & "OBX|13|ST|569^Calculated R Axis||13|degrees|||||F" & Chr(13) & "OBX|14|ST|560^T Axis||34|degrees|||||F" & Chr(13) & "OBX|15|ST|570^Calculated T Axis||34|degrees|||||F" & Chr(13) & _
                "OBX|16|ST|561^QRS Count||13|beats|||||F" & Chr(13) & "OBX|17|ST|562^Q Onset||226|ms|||||F" & Chr(13) & "OBX|18|ST|563^Q Offset||269|ms|||||F" & Chr(13) & "OBX|19|ST|564^P Onset||151|ms|||||F" & Chr(13) & _
                "OBX|20|ST|565^P Offset||202|ms|||||F" & Chr(13) & "OBX|21|ST|566^T Offset||423|ms|||||F" & Chr(13) & "OBX|22|ST|575^QTC Fredericia||432|ms|||||F" & Chr(13) & "OBX|23|ST|576^QTC Framingham||432|ms|||||F" & Chr(13) & _
                "OBX|24|ST|578^QTC RR|||ms|||||F" & Chr(13) & "OBX|25|ST|15006^MAC1200 PP Interval||759|ms|||||F" & Chr(13) & "OBX|26|ST|15002^MAC1200 RR Interval||758|ms|||||F" & Chr(13) & "OBX|27|RP|MUSEWebURL||http://ELECTROCARDDIOG/musescripts/museweb.dll?RetrieveTestByDateTime?PatientID=0000001202220054\T\Date=22-02-2012\T\Time=10%3a17%3a43%3a00\T\TestType=ECG\T\Site=1\T\OutputType=PDF\T\Ext=PDF||||||F" & Chr(13) & _
                "OBX|28|ST|15601^PONA^P onset amplitude|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|9^34^-10^14^14^24^24^29^24^-25^-5^29|uV|||||F|" & Chr(13) & "OBX|29|ST|15602^PA^P Peak Amplitude|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|78^131^48^48^63^68^78^83^83^-107^-29^107|uV|||||F|" & Chr(13) & "OBX|30|ST|15603^PD^P Duration|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|102^102^54^102^102^102^102^102^102^102^62^102|msec|||||F|" & Chr(13) & _
                "OBX|31|ST|15604^bmPAR^P Area|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|166^386^81^142^199^234^250^225^220^-282^-37^309|uV*msec|||||F|" & Chr(13) & "OBX|32|ST|15605^bmPI^P Peak Time|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|66^50^32^28^40^40^52^56^40^66^36^50|msec|||||F|" & Chr(13) & "OBX|33|ST|15606^PPA^P prime Amplitude|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|0^0^-24^0^0^0^0^0^0^0^0^0|uV|||||F|" & Chr(13) & _
                "OBX|34|ST|15607^PPD^P prime Duration|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|0^0^48^0^0^0^0^0^0^0^0^0|msec|||||F|" & Chr(13) & "OBX|35|ST|15608^bmPPAR^P prime Area|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|0^0^-19^0^0^0^0^0^0^0^0^0|uV*msec|||||F|" & Chr(13) & "OBX|36|ST|15609^bmPPI^P prime Peak Time|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|0^0^60^0^0^0^0^0^0^0^0^0|msec|||||F|" & Chr(13) & "OBX|37|ST|15610^QA^Q Peak Amplitude|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|43^0^0^0^0^0^0^0^0^727^63^0|uV|||||F|" & Chr(13) & _
                "OBX|38|ST|15611^QD^Q Duration|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|16^0^0^0^0^0^0^0^0^58^18^0|msec|||||F|" & Chr(13) & "OBX|39|ST|15612^bmQAR^Q Area|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|22^0^0^0^0^0^0^0^0^754^33^0|uV*msec|||||F|" & Chr(13) & "OBX|40|ST|15613^bmQI^Q Peak Time|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|10^0^0^0^0^0^0^0^0^36^10^0|msec|||||F|" & Chr(13) & "OBX|41|ST|15614^RA^R Amplitude|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|844^625^380^1083^1191^1352^1318^991^83^39^541^253|uV|||||F|" & Chr(13) & _
                "OBX|42|ST|15615^RD^R Duration|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|70^57^34^43^47^49^51^57^21^12^68^56|msec|||||F|" & Chr(13) & "OBX|43|ST|15616^bmRAR^R Area|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|941^686^343^984^1250^1290^1237^996^51^15^662^287|uV*msec|||||F|" & Chr(13) & "OBX|44|ST|15617^bmRI^R Peak Time|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|36^38^26^30^32^36^36^36^10^64^36^42|msec|||||F|" & Chr(13) & "OBX|45|ST|15618^SA^S Amplitude|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|0^146^654^957^620^307^209^73^297^0^0^180|uV|||||F|" & Chr(13) & _
                "OBX|46|ST|15619^SD^S Duration|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|0^29^52^32^32^29^25^19^24^0^0^30|msec|||||F|" & Chr(13) & "OBX|47|ST|15620^bmSAR^S Area|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|0^97^727^743^455^245^162^42^209^0^0^149|uV*msec|||||F|" & Chr(13) & "OBX|48|ST|15621^bmSI^S Peak Time|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|0^64^54^54^54^54^58^62^30^0^0^64|msec|||||F|" & Chr(13) & "OBX|49|ST|15622^RPA^R prime Amplitude|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|0^0^0^0^0^0^0^0^122^0^0^0|uV|||||F|" & Chr(13) & "OBX|50|ST|15623^RPD^R prime Duration|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|0^0^0^0^0^0^0^0^10^0^0^0|msec|||||F|" & Chr(13) & _
                "OBX|51|ST|15624^bmRPAR^R prime Area|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|0^0^0^0^0^0^0^0^37^0^0^0|uV*msec|||||F|" & Chr(13) & "OBX|52|ST|15625^bmRPI^R prime Peak Time|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|0^0^0^0^0^0^0^0^50^0^0^0|msec|||||F|" & Chr(13) & "OBX|53|ST|15626^SPA^S prime Amplitude|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|0^0^0^0^0^0^0^0^214^0^0^0|uV|||||F|" & Chr(13) & "OBX|54|ST|15627^SPD^S prime Duration|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|0^0^0^0^0^0^0^0^31^0^0^0|msec|||||F|" & Chr(13) & "OBX|55|ST|15628^bmSPAR^SA prime Area|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|0^0^0^0^0^0^0^0^207^0^0^0|uV*msec|||||F|" & Chr(13) & _
                "OBX|56|ST|15629^bmSPI^S prime Peak Time|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|0^0^0^0^0^0^0^0^64^0^0^0|msec|||||F|" & Chr(13) & "OBX|57|ST|15630^STJ^ST at J Point|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|48^19^9^73^48^48^48^39^-30^-35^39^-5|uV|||||F|" & Chr(13) & "OBX|58|ST|15631^STM^ST at Mid ST|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|58^53^63^151^112^97^92^68^-5^-59^34^24|uV|||||F|" & Chr(13) & "OBX|59|ST|15632^STE^ST at End ST|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|87^87^92^239^190^170^146^102^0^-88^43^43|uV|||||F|" & Chr(13) & "OBX|60|ST|15633^MXSTA^Maximum ST Area|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|87^87^92^239^190^170^146^102^0^-35^43^43|uV|||||F|" & Chr(13) & _
                "OBX|61|ST|15634^MNSTA^Minimum ST Area|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|48^19^9^73^48^48^48^39^-30^-59^34^-5|uV|||||F|" & Chr(13) & "OBX|62|ST|15635^SPTA^Special T|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|142^137^44^210^185^210^220^181^-14^-229^79^74|uV|||||F|" & Chr(13) & "OBX|63|ST|15636^QRSA^QRS Balance|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|801^479^-274^126^571^1045^1109^918^-175^-688^478^73|uV|||||F|" & Chr(13) & "OBX|64|ST|15637^QRSDEF^QRS Deflection|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|887^771^1034^2040^1811^1659^1527^1064^419^766^604^433|uV|||||F|" & Chr(13) & "OBX|65|ST|15638^MAXRA^Maximum R Amplitude|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|844^625^380^1083^1191^1352^1318^991^122^39^541^253|uV|||||F|" & Chr(13) & _
                "OBX|66|ST|15639^MAXSA^Maximum S Amplitude|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|43^146^654^957^620^307^209^73^297^727^63^180|uV|||||F|" & Chr(13) & "OBX|67|ST|15640^TA^T Peak Amplitude|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|229^224^136^449^375^380^366^283^-14^-229^122^117|uV|||||F|" & Chr(13) & "OBX|68|ST|15641^TD^T Duration|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|214^214^214^214^214^214^214^214^116^214^174^214|msec|||||F|" & Chr(13) & "OBX|69|ST|15642^bmTAR^T Area|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|1403^1538^966^3156^2682^2630^2524^1884^-18^-1485^657^850|uV*msec|||||F|" & Chr(13) & "OBX|70|ST|15643^bmTI^T Peak Time|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|102^106^76^96^90^104^106^112^96^106^102^118|msec|||||F|" & Chr(13) & _
                "OBX|71|ST|15644^TPA^T prime Amplitude|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|0^0^0^0^0^0^0^0^53^0^0^0|uV|||||F|" & Chr(13) & "OBX|72|ST|15645^TPD^T prime Duration|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|0^0^0^0^0^0^0^0^98^0^0^0|msec|||||F|" & Chr(13) & "OBX|73|ST|15646^bmTPAR^T prime Area|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|0^0^0^0^0^0^0^0^149^0^0^0|uV*msec|||||F|" & Chr(13) & "OBX|74|ST|15647^bmTPI^T Prime Peak Time|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|0^0^0^0^0^0^0^0^160^0^0^0|msec|||||F|" & Chr(13) & "OBX|75|ST|15648^TEND^T End|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|4^24^19^53^48^48^43^43^19^-15^-5^24|uV|||||F|" & Chr(13) & "OBX|76|ST|15649^PAREA^P Wave Area (full)|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|166^386^69^141^199^234^250^225^220^-276^-27^303|uV*msec|||||F|" & Chr(13) & _
                "OBX|77|ST|15650^QRSAR^QRS Area|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|914^589^-383^263^803^1054^1085^958^-325^-752^620^132|uV*msec|||||F|" & Chr(13) & "OBX|78|ST|15651^TAREA^T Wave Area (full)|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|1403^1536^964^3150^2677^2625^2520^1880^133^-1470^635^835|uV*msec|||||F|" & Chr(13) & "OBX|79|ST|15652^QRSINT^QRS Intrinsicoid|I^II^V1^V2^V3^V4^V5^V6^III^AVR^AVL^AVF|36^38^54^30^32^36^36^36^30^36^36^42|msec|||||F|" & Chr(13) & "OBX|80|ST|15653^PON^P Onset||151||||||F|" & Chr(13) & "OBX|81|ST|15654^POFF^P Offset||202||||||F|" & Chr(13) & "OBX|82|ST|15655^QRSON^QRS Onset||226||||||F|" & Chr(13) & "OBX|83|ST|15656^QRSOFF^QRS Offset||269||||||F|" & Chr(13) & "OBX|84|ST|15657^TON^T Onset||316||||||F|" & Chr(13) & _
                "OBX|85|ST|15658^TOFF^T Offset||423||||||F|" & Chr(13) & "OBX|86|TX|208.0^Diagnosis||窦性心律 ||||||F" & Chr(13) & "OBX|87|TX|208.1^Coded Diagnosis||窦性心律||||||F" & Chr(13) & "OBX|88|FT|ECGMEASANDDIAG||Test Reason : ~Blood Pressure : ***/*** mmHG~Vent. Rate : 079 BPM     Atrial Rate : 079 BPM~   P-R Int : 150 ms          QRS Dur : 086 ms~    QT Int : 394 ms       P-R-T Axes : 065 013 034 degrees~   QTc Int : 451 ms~~窦性心律 ~~Referred By:             Overread By: 勇娟 郭||||||D|" & Chr(13) & _
                Chr(28) & Chr(13)


'strData = Chr$(11) & "MSH|^~\&|SCM|AGH|MUSE|SITE0001|20120222094137||ORM^O01||P|2.4" & Chr(13) & _
'                "PID|1||病人号|患者第二编号|名^姓||19591105|M|中文名|C|地址1^地址2^市^省^邮编^国籍||电话1|电话2|||Protestant|150035926198|232-654-9988" & Chr(13) & _
'                "PV1|1|INPAT|HIS科室^房间号^床位号^^^^^^第二科室|X|||会诊医生ID^SHARIF^SAIMA|转诊医师ID^Dosch^Justin|主治医师ID^Last Name^First Name|HON||||EO|8||接诊医师ID^接诊医师名^接诊医师姓|S|就诊号|||||||||||||||||UNK|||Fac|||||201202220941|201203050941|||||第二就诊号" & Chr(13) & _
'                "ORC|NW|医嘱号^SCM|fillers number||||^^^201202230941^^Routine|parent num|201202220941|Place医师ID^Martin^Tom||开单医生ORDID^Dosch^Justin" & Chr(13) & _
'                "OBR|1|||ECGGDT^12 Lead ECG^60200151^^12 Lead ECG||20070617235953|20070617235953|||||明尼苏达码||||535724^Dosch Justin|||||||||||^^^^^Routine||||^检查原因 (786.50)" & Chr(13) & _
'                "NTE|1||医嘱注解|" & Chr(13) & _
'                "OBX|1||||^181|" & Chr(13) & _
'                "OBX|2||||^80|" & Chr(13) & _
'                "DG1|1||第一结论|Admit Diagnosis|" & Chr(13) & _
'                "DG1|2||Secondary diagnosis||" & Chr(13) & _
'                "DG1|3||Tertiary diagnosis||" & Chr(13)
                
                
                
'strData = Chr(11) & "MSH|^~\&|SCM|ZLHIS|HIS001|SITE0001|20120222094137||ORM^O01||P|2.4" & Chr(13) & _
'                "PID|1||病人号5|患者第二编号|名^姓||19591105|M|中联测试|C|地址1^地址2^市^省^邮编^国籍||电话1|电话2|||Protestant|150035926198|232-654-9988" & Chr(13) & _
'                "PV1|1|INPAT|HIS科室^房间号^床位号^^^^^^第二科室|X|||会诊医生ID^SHARIF^SAIMA|转诊医师ID^Dosch^Justin|主治医师ID^Last Name^First Name|HON||||EO|8||接诊医师ID^接诊医师名^接诊医师姓|S|就诊号|||||||||||||||||UNK|||Fac|||||201202220941|201203050941|||||第二就诊号" & Chr(13) & _
'                "ORC|NW|123456^SCM|fillers number||||^^^201202230941^^Routine|parent num|201202220941|Place医师ID^Martin^Tom||开单医生ORDID^Dosch^Justin" & Chr(13) & _
'                "OBR|1|||ECGGDT^12 Lead ECG^60200151^^12 Lead ECG||20070617235953|20070617235953|||||明尼苏达码||||535724^Dosch Justin|||||||||||^^^^^Routine||||^检查原因 (786.50)" & Chr(13) & _
'                "NTE|1||医嘱注解|" & Chr(13) & _
'                "OBX|1||||^181|" & Chr(13) & _
'                "OBX|2||||^80|" & Chr(13) & _
'                "DG1|1||第一结论|Admit Diagnosis|" & Chr(13) & _
'                "DG1|2||Secondary diagnosis||" & Chr(13) & _
'                "DG1|3||Tertiary diagnosis||" & Chr(13) & _
'                Chr(28) & Chr(13)
                
'strData = Chr(11) & "MSH|^~\&|ZLHIS|HIS001|MUSE|SITE0001|20120222094137||ORM^O01||P|2.4" & Chr(13) & _
'                    "PID|1||病人号|患者第二编号|名^姓||19591105|M|中文名|C|地址1^地址2^市^省^邮编^国籍||电话1|电话2|||Protestant|150035926198|232-654-9988" & Chr(13) & _
'                    "PV1|1|INPAT|HIS科室^房间号^床位号^^^^^^第二科室|X|||会诊医生ID^SHARIF^SAIMA|转诊医师ID^Dosch^Justin|主治医师ID^Last Name^First Name|HON||||EO|8||接诊医师ID^接诊医师名^接诊医师姓|S|就诊号|||||||||||||||||UNK|||Fac|||||201202220941|201203050941|||||第二就诊号" & Chr(13) & _
'                    "ORC|CA|医嘱号^SCM|fillers number||||^^^201202230941^^Routine|parent num|201202220941|Place医师ID^Martin^Tom||开单医生ORDID^Dosch^Justin" & Chr(13) & _
'                    "OBR|1|||ECGGDT^12 Lead ECG^60200151^^12 Lead ECG||20070617235953|20070617235953|||||明尼苏达码||||535724^Dosch Justin|||||||||||^^^^^Routine||||^检查原因 (786.50)" & Chr(13) & _
'                    "ZEX||adt extra one^adt extra two^adt extra three^adt extra four|visit extra one^visit extra two^visit extra three^visit extra four|order extra one^order extra two^order extra three^order extra four|medications|||||promt one^额外问题1|promt two^data two|promt 3^data three|num^data four|" & Chr(13) & _
'                    Chr(28) & Chr(13)
                
     
     Me.Winsock1.SendData strData
     Me.txtInfo.Text = Me.txtInfo.Text & "发送数据：" & vbCrLf & strData & vbCrLf
End Sub

Private Sub Winsock1_Connect()
    Me.txtInfo.Text = Me.txtInfo.Text & "连接成功，IP：" & Me.txtIP.Text & " PORT:" & Me.txtPort.Text & vbCrLf
    Call Command3_Click
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Me.Winsock1.GetData strData, vbString
    Me.txtInfo.Text = Me.txtInfo.Text & "接收到ACK:" & vbCrLf & strData & vbCrLf
End Sub

