VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_LG_PABER_I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'ÒÇÆ÷:ÑªĞ¡°å¾Û¼¯ÄıÑªÒò×Ó·ÖÎöÒÇ
'ĞÍºÅ:LG-PABER-I
'³§ÉÌ:
'×÷Õß:Ôø³¬

Public Function GetStartCmd() As String
    'Éè±¸±ê×¼½Ó¿Ú
    '¹¦ÄÜ£º»ñÈ¡Éè±¸¿ªÊ¼·¢ËÍÃüÁî
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    'Éè±¸±ê×¼½Ó¿Ú
    '¹¦ÄÜ£º»ñÈ¡Éè±¸Í£Ö¹·¢ËÍÃüÁî
    GetEndCmd = ""
End Function

Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
    strCmd As String)
    'Éè±¸±ê×¼½Ó¿Ú
    '¹¦ÄÜ£º½âÎöÊı¾İ
    '²ÎÊı£º
    'strSource£ºÒª½âÎöµÄÔ­Ê¼´®
    'strResult£º·µ»ØµÄ¼ìÑé½á¹û(¸÷ÒÇÆ÷½âÎö³ÌĞò±ØĞë°´ÒÔÏÂ±ê×¼×éÖ¯½á¹û)
    '   Ã¿×é¼ìÑé½á¹ûÒÔ||·Ö¸ô,ÔªËØÖ®¼äÒÔ|·Ö¸ô
    '   µÚ0¸öÔªËØ£º¼ìÑéÊ±¼ä
    '   µÚ1¸öÔªËØ£ºÑù±¾ĞòºÅ
    '   µÚ2¸öÔªËØ£º¼ìÑéÈË
    '   µÚ3¸öÔªËØ£º±ê±¾
    '   µÚ4¸öÔªËØ£ºÊÇ·ñÖÊ¿ØÆ·
    '   ´ÓµÚ5¸öÔªËØ¿ªÊ¼Îª¼ìÑé½á¹û£¬Ã¿2¸öÔªËØ±íÊ¾Ò»¸ö¼ìÑéÏîÄ¿¡£
    '       Èç£ºµÚ5i¸öÔªËØÎª¼ìÑéÏîÄ¿£¬µÚ5i+1¸öÔªËØÎª¼ìÑé½á¹û
    'strReserved£º×îºó²»ÄÜÍê³É½âÎöµÄÔ­Ê¼´®£¬ĞèÒª·µ»Ø×÷ºóĞø´¦Àí
    'strCmd£ºÈç¹ûĞèÒª£¬¿É·µ»ØÏòÉè±¸·¢ËÍµÄÃüÁî
    
    Dim iStrPoint As Integer '×Ö·ûÖ¸Õë
    Dim strBuffer As String, strData As String
    Dim iCurrGrpEntry As Integer 'µ±Ç°Õâ×é½á¹ûµÄÈë¿Ú×Ö·ûÎ»ÖÃ
    
    Dim varIn As Variant '´æ·ÅÊäÈëÊı¾İ
    Dim StrTime As String
    
    Dim lngSector As Long, lngCup As Long, strSampleID As String, strTmpID As String
    Dim strLine As String
    Dim str±ê±¾ºÅ As String, strÈÕÆÚ As String, str¼ìÑéÈË±ê±¾ÖÊ¿Ø As String, strÖ¸±ê As String
    Dim aLine() As String
    Dim strDate As String
    Dim intLoop As Integer
    Dim aItem() As String
    Dim strNO As String
    Dim strNow As String
    Dim strItem As String
    Dim intCount As Integer
    Const StrAllitem As String = ",ADP,PT,TT,APTT,FIB,"
    
    On Error GoTo errHandle
    
    '³õÊ¼»¯·µ»ØÖµ
    strReserved = "": strResult = "": strCmd = ""
    
    'Ã»ÓĞ¿ªÊ¼ºÍ½áÊø×Ö·ûÊ±ÍË³ö
    If InStr(strSource, "@c") = 0 Or InStr(strSource, Chr(&HD) & Chr(&HD) & Chr(&HD) & Chr(&HD) & Chr(&HD)) = 0 Then
        strReserved = strSource
        Exit Sub
    End If
    
    strNow = Mid(strSource, InStr(strSource, "@c") + 5, InStr(strSource, Chr(&HD) & Chr(&HD) & Chr(&HD) & Chr(&HD) & Chr(&HD)) + 5)
    
    If InStr(strNow, "Meter") > 0 Then
        strNow = Mid(strNow, InStr(strNow, "Meter") + 6)
    End If
    
    aLine = Split(strNow, vbCr)
    
    strDate = Replace(Mid(aLine(0), 1, 10), ".", "-") & " " & Time
    
    For intLoop = 3 To UBound(aLine)
        strLine = aLine(intLoop)
        strLine = Replace(strLine, Chr(30), " ")
        strLine = Replace(strLine, Chr(0), " ")
        strLine = Replace(strLine, Chr(9), " ")
        
        strLine = DelSapce(strLine)
        aItem = Split(strLine, " ")
        
        If InStr(strLine, "") > 0 Or InStr(strLine, Chr(&H14) & Chr(&H1E)) > 0 Then
            strNO = aItem(1)
        End If
       
        If UBound(aItem) > 1 Then
        
            If InStr(strLine, " ADP ") Then
                strItem = "ADP"
            End If
            
            If InStr(strLine, " PT ") Then
                strItem = "PT"
                strData = strData & "||" & strDate & "|" & strNO & "| |ÑªÇå|0|PT|" & aItem(UBound(aItem) - 2)
            End If
            
            If InStr(strLine, " TT ") Then
                strData = strData & "||" & strDate & "|" & strNO & "| |ÑªÇå|0|TT|" & aItem(UBound(aItem) - 1)
                strItem = ""
            End If
            
            If InStr(strLine, " APTT ") Then
                strData = strData & "||" & strDate & "|" & strNO & "| |ÑªÇå|0|APTT|" & aItem(UBound(aItem) - 1)
                strItem = ""
            End If
            
            If InStr(strLine, " FIB ") Then
                strItem = ""
                strData = strData & "||" & strDate & "|" & strNO & "| |ÑªÇå|0|FIB|" & aItem(UBound(aItem) - 1)
            End If
            
            If strItem = "ADP" Then
                If aItem(1) = "180" Then
                    strData = strData & "||" & strDate & "|" & strNO & "| |ÑªÇå|0|ADP|" & aItem(UBound(aItem) - 1)
                End If
            End If
            
            If strItem = "PT" Then
                If aItem(2) = "INR" Then
                    strData = strData & "||" & strDate & "|" & strNO & "| |ÑªÇå|0|INR|" & aItem(UBound(aItem) - 1)
                End If
            End If
        End If
        
    Next
    
    strResult = Mid(strData, 3)

    Exit Sub
    

errHandle:
    WriteLog "lg_Paber_i,½âÂë´íÎó", Err.Description, ""

End Sub

Public Function SendSample(ByVal strSample As String) As String
    '½âÎöĞèÒªÏòÒÇÆ÷·¢ËÍµÄ±ê±¾ĞÅÏ¢
    '   ÔªËØÖ®¼äÒÔ|·Ö¸ô
    '   µÚ0¸öÔªËØ£º¼ìÑéÊ±¼ä
    '   µÚ1¸öÔªËØ£ºÑù±¾ĞòºÅ
    '   µÚ2¸öÔªËØ£º¼ìÑéÈË
    '   µÚ3¸öÔªËØ£º±ê±¾
    '   µÚ4¸öÔªËØ£ºÊÇ·ñÖÊ¿ØÆ·
    '   µÚ5¡«9ÔªËØ£ºÏµÍ³±£Áô
    '   ´ÓµÚ10¸öÔªËØ¿ªÊ¼ÎªĞèÒªµÄ¼ìÑéÏîÄ¿¡£
    SendSample = ""
End Function












