VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'################################################################################################################
'##模 块 名：cMedia.cls
'##创 建 人：吴庆伟
'##日    期：2004年11月29日
'##修 改 人：
'##日    期：
'##描    述：多媒体类（用于播放声音和视频）
'##
'##
'##   使用方法：
'##   Dim Multimedia As New cMedia    '初始化（请设为全局变量）
'##
'##   Multimedia.Wait = False         '是否等待播放完毕
'##   Multimedia.mmOpen strFileName   '打开媒体
'##   PBar1.Max = Multimedia.Length   '媒体播放长度
'##   Multimedia.mmPlay               '开始播放
'##
'##   类型可以为："WaveAudio (*.wav)|*.wav|Midi (*.mid)|*.mid|Video files (*.avi)|*.avi"
'##
'################################################################################################################

Option Explicit

'-----------------------------------------------------
'   Name    :   MMedia.cls
'   Author  :   Peter Wright, For BG2VB4 & BG2VB5
'
'   Notes   :   A multimedia class, which when turned
'           :   into an object lets you load and play
'           :   multimedia files, such as sound and
'           :   video.
'-----------------------------------------------------

' -=-=-=- PROPERTIES -=-=-=-
' Filename      Determines the name of the current file
' Length        The length of the file (Read Only)
' Position      The current position through the file
' Status        The current status of the object (Read Only)
' Wait          True/False...tells VB to wait until play done

' -=-=-=- METHODS -=-=-=-=-
' mmOpen <Filename>   Opens the requested filename
' mmClose             Closes the current file
' mmPause             Pauses playback of the current file
' mmStop              Stops playback ready for closedown
' mmSeek <Position>   Seeks to a position in the file
' mmPlay              Plays the open file

'-------------------------------------------------------------
' NOTES
' -----
'
' Open a file, then play it. Pause it in response to a request
' from the user. Stop if you intend to seek to the start and
' play again. Close when you no longer want to play the file
'--------------------------------------------------------------

Private sAlias As String        ' Used internally to give an alias name to
                          ' the multimedia resource

Private sFilename As String     ' Holds the filename internally
Private nLength As Single       ' Holds the length of the filename
                          ' internally
Private nPosition As Single     ' Holds the current position internally
Private sStatus As String       ' Holds the current status as a string
Private bWait As Boolean        ' Determines if VB should wait until play
                        ' is complete before returning.

'------------ API DECLARATIONS -------------
'note that this is all one code line:
Private Declare Function mciSendString Lib "winmm.dll" _
   Alias "mciSendStringA" (ByVal lpstrCommand As String, _
   ByVal lpstrReturnString As String, ByVal uReturnLength As Long, _
   ByVal hwndCallback As Long) As Long

'################################################################################################################
'## 用途：  打开媒体文件
'##
'## 参数：  sTheFile    :文件名
'################################################################################################################
Public Sub mmOpen(ByVal sTheFile As String)

    ' Declare a variable to hold the value returned by mciSendString
    Dim nReturn As Long
    
    ' Declare a string variable to hold the file type
    Dim sType As String

    ' Opens the specified multimedia file, and closes any
    ' other that may be open
    If sAlias <> "" Then
        mmClose
    End If
    
    ' Determine the type of file from the file extension
    Select Case UCase$(Mid(sTheFile, Len(sTheFile) - 2))
       Case "WAV"
          sType = "Waveaudio"
       Case "AVI"
          sType = "AviVideo"
       Case "MID"
          sType = "Sequencer"
       Case Else
          ' If the file extension is not known then exit the subroutine
          Exit Sub
    End Select
    sAlias = Mid(sTheFile, Len(sTheFile) - 2) & Minute(Now)

    ' At this point there is no file open, and we have determined the
    ' file type. Now would be a good time to open the new file.
    ' Note: if the name contains a space we have to enclose it in quotes
    If InStr(sTheFile, " ") Then sTheFile = Chr(34) & sTheFile & Chr(34)
    nReturn = mciSendString("Open " & sTheFile & " ALIAS " & sAlias _
            & " TYPE " & sType & " wait", "", 0, 0)
End Sub

'################################################################################################################
'## 用途：  关闭
'################################################################################################################
Public Sub mmClose()
    ' Closes the currently opened multimedia file

    ' Declare a variable to hold the return value from the mciSendString
    ' command
    Dim nReturn As Long

    ' If there is no file currently open then exit the subroutine
    If sAlias = "" Then Exit Sub
    
    nReturn = mciSendString("Close " & sAlias, "", 0, 0)
    sAlias = ""
    sFilename = ""
    
End Sub

'################################################################################################################
'## 用途：  暂停播放
'################################################################################################################
Public Sub mmPause()
    ' Pause playback of the file

    ' Declare a variable to hold the return value from the mciSendString
    ' command
    Dim nReturn As Long
    
    ' If there is no file currently open then exit the subroutine
    If sAlias = "" Then Exit Sub
    
    nReturn = mciSendString("Pause " & sAlias, "", 0, 0)

End Sub

'################################################################################################################
'## 用途：  播放
'################################################################################################################
Public Sub mmPlay()
    ' Plays the currently open file, from the current position

    ' Declare a variable to hold the return value from the mciSendString
    ' command
    Dim nReturn As Long
    
    ' If there is no file currently open, then exit the routine
    If sAlias = "" Then Exit Sub
    
    ' Now play the file
    If bWait Then
        nReturn = mciSendString("Play " & sAlias & " wait", "", 0, 0)
    Else
        nReturn = mciSendString("Play " & sAlias, "", 0, 0)
    End If
End Sub

'################################################################################################################
'## 用途：  停止播放
'################################################################################################################
Public Sub mmStop()
    ' Stop using a file totally, be it playing or whatever

    ' Declare a variable to hold the return value from mciSendString
    Dim nReturn As Long
    
    ' If there is no file currently open then exit the subroutine
    If sAlias = "" Then Exit Sub
    
    nReturn = mciSendString("Stop " & sAlias, "", 0, 0)
    
End Sub

'################################################################################################################
'## 用途：  跳转到指定位置
'##
'## 参数：  nPosition   :需要跳到的位置
'################################################################################################################
Public Sub mmSeek(ByVal nPosition As Single)
    ' Seeks to a specific position within the file

    ' Declare a variable to hold the return value from the mciSendString
    ' function
    Dim nReturn As Long
    
    nReturn = mciSendString("Seek " & sAlias & " to " & nPosition, "", 0, 0)

End Sub

'################################################################################################################
'## 用途：  返回当前的文件名
'################################################################################################################
Property Get Filename() As String
' Routine to return a value when the programmer asks the
' object for the value of its Filename property
    Filename = sFilename
End Property

'################################################################################################################
'## 用途：  设置当前的文件名
'################################################################################################################
Property Let Filename(ByVal sTheFile As String)
' Routine to set the value of the filename property, should the programmer
' wish to do so. This implies that the programmer actually wants to open
' a file as well so control is passed to the mmOpen routine
   mmOpen sTheFile
End Property

'################################################################################################################
'## 用途：  返回设备 Wait 属性的值
'################################################################################################################
Property Get Wait() As Boolean
' Routine to return the value of the object's wait property.
   Wait = bWait
End Property

'################################################################################################################
'## 用途：  设置设备 Wait 属性的值
'################################################################################################################
Property Let Wait(bWaitValue As Boolean)
' Routine to set the value of the object's wait property
   bWait = bWaitValue
End Property

'################################################################################################################
'## 用途：  返回文件播放长度
'################################################################################################################
Property Get Length() As Single
   ' Routine to return the length of the currently opened multimedia file

   ' Declare a variable to hold the return value from the mciSendString
   Dim nReturn As Long, nLength As Integer

   ' Declare a string to hold the returned length from the mci Status call
   Dim sLength As String * 255
    
   ' If there is no file open then return 0
   If sAlias = "" Then
      Length = 0
      Exit Property
   End If

  nReturn = mciSendString("Status " & sAlias & " length", sLength, 255, 0)
  nLength = InStr(sLength, Chr$(0))
  Length = Val(Left$(sLength, nLength - 1))
End Property

'################################################################################################################
'## 用途：  设置播放的当前位置
'################################################################################################################
Property Let Position(ByVal nPosition As Single)
' Sets the Position property effectively by seeking
    mmSeek nPosition
End Property

'################################################################################################################
'## 用途：  返回播放的当前位置
'################################################################################################################
Property Get Position() As Single
   ' Returns the current position in the file
    
   ' Declare a variable to hold the return value from mciSendString
   Dim nReturn As Integer, nLength As Integer
    
   ' Declare a variable to hold the position returned
   ' by the mci Status position command
   Dim sPosition As String * 255

   ' If there is no file currently opened then exit the subroutine
   If sAlias = "" Then Exit Property
    
   ' Get the position and return
   nReturn = mciSendString("Status " & sAlias & " position", sPosition, 255, 0)
   nLength = InStr(sPosition, Chr$(0))
   Position = Val(Left$(sPosition, nLength - 1))

End Property

'################################################################################################################
'## 用途：  返回播放状态字符串
'################################################################################################################
Property Get Status() As String
   ' Returns the playback/record status of the current file

   ' Declare a variable to hold the return value from mciSendString
   Dim nReturn As Integer, nLength As Integer
    
   ' Declare a variable to hold the return string from mciSendString
   Dim sStatus As String * 255
    
   ' If there is no file currently opened, then exit the subroutine
   If sAlias = "" Then Exit Property

   nReturn = mciSendString("Status " & sAlias & " mode", sStatus, 255, 0)
    
   nLength = InStr(sStatus, Chr$(0))
   Status = Left$(sStatus, nLength - 1)
End Property






