VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' CLASS_MODULE:   VBA Logger - allows 'log4VBA' style Logging in VBA
'''                 - please see the 'Logging' Module for Usage: the Logging Module
'''                   automatically creates a 'Logger' instance and provides additional
'''                   Features
'''
'''                 - use Macro 'Test' from 'TestLogging' for testing and as an example
'''
''' Date            Developer               Action
''' --------------------------------------------------------------------------
''' 28/08/08        Christian Bolterauer    Created
'''

Public cLogbuffer As Logbuffer

Public iLogLevel As Integer
Public bUseLogPrefix As Boolean
Public bConsole As Boolean
Public bBuffer As Boolean
Public bToLogFile As Boolean
Public LogFilePath As String
Public TraceFilePath As String
Public bDelLogFileAtSetup As Boolean
Public PropsFileName As String

'ModulName
Public ModulName As String

'Define log levels
Public Enum LogLEVEL
  DISABLED = 0
  BASIC = 1 'like Debug.Print
  FATAL = 2
  WARN = 3
  INFO = 4
  FINE = 5
  FINER = 6
  FINEST = 7
  ALL = 8
End Enum

' The defaults
Const DEFAULT_LOG_LEVEL% = LogLEVEL.INFO
Const DEFAULT_LOG_Console = True
Const DEFAULT_LOG_Buffer = False
Const DEFAULT_LOG_FILE = False
Const DEFAULT_PROPSFILE_NAME = "vba_log.properties"
'


'Class Konstructor
Private Sub Class_Initialize()
   On Error GoTo Errhandler:
   Dim localpath As String
   Set cLogbuffer = New Logbuffer
   bUseLogPrefix = True
   bDelLogFileAtSetup = True
   
   'default
   ModulName = ""
   
   'set default location of props file to directory of this Logger and add default name
   localpath = getParentFolder(Application.VBE.ActiveVBProject.Filename) 'set path to location of file containing this Logger
   PropsFileName = localpath & "\" & DEFAULT_PROPSFILE_NAME
   
   'make sure defaults are set
   Call setLoggigParams(DEFAULT_LOG_LEVEL, DEFAULT_LOG_Console, DEFAULT_LOG_Buffer, DEFAULT_LOG_FILE)
   'set default log file path
   LogFilePath = localpath & "\" & "vba_logger.log"
   TraceFilePath = localpath & "\" & "vba_trace.log"
   'check if params can be set from a properties file and overwrite defaults if available
   Call getLogParamsFromFile
   'set log file
   Call setLogFile(LogFilePath, False)

Exit Sub

Errhandler:
  Debug.Print "Error in Logger.Class_Initialize & "; " & Err.Number & "; " & Err.Description"
      
End Sub

'set logging parameters
Public Sub setLoggigParams(level As Integer, toConsole As Boolean, toBuffer As Boolean, toLogFile As Boolean, Optional deleteExistingLogFile)
     Dim delLogfile As Boolean
     
     iLogLevel = level
     bConsole = toConsole
     bBuffer = toBuffer
     bToLogFile = toLogFile
     
     If IsMissing(deleteExistingLogFile) Then
        delLogfile = False
     Else
        delLogfile = deleteExistingLogFile
     End If
     ' delete currently set Logfile if set
     If (delLogfile) Then deleteLogFile
     
     
End Sub

'The main log procedure
Public Sub log(sLogText As String, level As LogLEVEL, Optional slogpoint As String)
    
    If (Me.iLogLevel > LogLEVEL.DISABLED And Me.iLogLevel >= level) Then
        If IsMissing(slogpoint) Then
           Call WriteLog(sLogText, level, "")
        Else
           Call WriteLog(sLogText, level, slogpoint)
        End If
    End If
End Sub

Public Sub logBASIC(sLogText As String, Optional slogpoint As String)
   Call Me.log(sLogText, LogLEVEL.BASIC, slogpoint)
End Sub
Public Sub logINFO(sLogText As String, Optional slogpoint As String)
   Call Me.log(sLogText, LogLEVEL.INFO, slogpoint)
End Sub
Public Sub logWARN(sLogText As String, Optional slogpoint As String)
   Call Me.log(sLogText, LogLEVEL.WARN, slogpoint)
End Sub
Public Sub logFATAL(sLogText As String, Optional slogpoint As String)
   Call Me.log(sLogText, LogLEVEL.FATAL, slogpoint)
End Sub
Public Sub logFINE(sLogText As String, Optional slogpoint As String)
   Call Me.log(sLogText, LogLEVEL.FINE, slogpoint)
End Sub
Public Sub logFINER(sLogText As String, Optional slogpoint As String)
   Call Me.log(sLogText, LogLEVEL.FINER, slogpoint)
End Sub
Public Sub logFINEST(sLogText As String, Optional slogpoint As String)
   Call Me.log(sLogText, LogLEVEL.FINEST, slogpoint)
End Sub

Private Sub WriteLog(sLogText, level As LogLEVEL, slogpoint As String)
  
  Dim LogMessage As String
  Dim sDateTime As String
  Dim sLogPrefix As String
  
  LogMessage = getLogPrefix(level, slogpoint) & sLogText

  ' write to console
  If Me.bConsole Then Debug.Print (LogMessage)
  ' write to Buffer
  If Me.bBuffer Then cLogbuffer.addline (LogMessage)
  ' write to file
  If Me.bToLogFile Then writeToLogFile (LogMessage)

End Sub

' get LogLevelName for Integer value
Public Function getLogLevelName(level As Integer)
    Dim myLevelName As String
    
    Select Case level
    Case LogLEVEL.DISABLED:
       myLevelName = "DISABLED"
    Case LogLEVEL.BASIC:
       myLevelName = "BASIC"
    Case LogLEVEL.INFO:
       myLevelName = "INFO:"
    Case LogLEVEL.WARN:
       myLevelName = "WARN:"
    Case LogLEVEL.FATAL:
       myLevelName = "FATAL:"
    Case LogLEVEL.FINE:
       myLevelName = "FINE:"
    Case LogLEVEL.FINER:
       myLevelName = "FINER:"
    Case LogLEVEL.FINEST:
       myLevelName = "FINEST:"
    Case LogLEVEL.ALL:
       myLevelName = "ALL:"
       
    Case Else
       myLevelName = "Level is not defined:"
  End Select

  getLogLevelName = myLevelName

End Function

Private Function getLogPrefix(level As LogLEVEL, logpoint As String)
  Dim sDateTime As String
  Dim myLevelPrefix As String
  Dim mySubModul As String
  Dim iLevel As Integer
  
  If Not (bUseLogPrefix) Or level = LogLEVEL.BASIC Then   'when level = LogLEVEL.BASIC no prefix to simulate Debug.Print
     getLogPrefix = ""
     Exit Function
  End If
  
  iLevel = level ' to Integer
  myLevelPrefix = getLogLevelName(iLevel)
  
  If (Len(Me.ModulName) > 0 And Len(logpoint) > 0) Then
     mySubModul = "[" & Me.ModulName & "::" & logpoint & "]"
  ElseIf (Len(logpoint) > 0) Then
     mySubModul = "[" & logpoint & "]"
  ElseIf (Len(Me.ModulName) > 0) Then
     mySubModul = "[" & Me.ModulName & "]"
  Else
     mySubModul = ""
  End If
  
  sDateTime = CStr(Now())
  'ToDo provide different output styles ..
  'getLogPrefix = myLevelPrefix & " (" & sDateTime & ") - "
  getLogPrefix = "(" & sDateTime & ")" & mySubModul & "-" & myLevelPrefix & "  "
  
End Function

Private Sub writeToLogFile(logmsg As String)
    On Error GoTo Errhandler:
    If Len(Me.LogFilePath) = 0 Then
       Debug.Print "Error: Log file path is empty."
       Exit Sub
    End If
       
    Dim FileNum As Integer
    FileNum = FreeFile ' next file number
    Open Me.LogFilePath For Append As #FileNum ' creates the file if it doesn't exist
    Print #FileNum, logmsg ' write information at the end of the text file
    Close #FileNum ' close the file
Exit Sub

Errhandler:
    Debug.Print "Error writing to Logfile: " & Me.LogFilePath & " " & Err.Number & " " & Err.Description
    
End Sub

Public Sub writeLogBufferToTraceFile(Optional myfilePath As String)
    On Error GoTo Errhandler:
    Dim mytracefile As String
    
    If Len(myfilePath) = 0 Then
        mytracefile = Me.TraceFilePath
    Else
        mytracefile = myfilePath
    End If
    If Len(mytracefile) = 0 Then
       Me.logFATAL "Error: Trace file path is empty."
       Exit Sub
    End If
    'write to trace file
    Me.cLogbuffer.writeLogBufferToTraceFile (mytracefile)
Exit Sub

Errhandler:
    Debug.Print "Error writing to Tracefile: " & mytracefile & " " & Err.Number & " " & Err.Description
End Sub

Private Sub readPropertiesFile(path As String)
    
    On Error GoTo Errhandler:
    Dim txtline As String
    Dim para() As String
    Dim mymsg As String
    
    If Len(path) = 0 Then GoTo Errhandler
    Open path For Input As #1    ' open file
    Do While Not EOF(1)          ' Loop until end of file
      Line Input #1, txtline     ' read line
      
      'Debug.Print txtline 'test
      para = readParameter(txtline)
    
      If Len(para(0)) = 0 Then
        'continue
      ElseIf ("LOG_LEVEL" = UCase(para(0))) Then
        Call setLogLevel(para(1))
      ElseIf ("LOG_TO_CONSOLE" = UCase(para(0))) Then
        bConsole = valIsTrue(para(1))
      ElseIf ("LOG_TO_BUFFER" = UCase(para(0))) Then
        bBuffer = valIsTrue(para(1))
      ElseIf ("LOG_TO_FILE" = UCase(para(0))) Then
        bToLogFile = valIsTrue(para(1))
      ElseIf ("LOG_FILE_PATH" = UCase(para(0))) Then
        Me.LogFilePath = para(1)
      End If
    Loop
    Close #1
    'show settings
    mymsg = "Logging with logLevel=" & getLogLevelName(iLogLevel) & " toConsole=" & bConsole & " toBuffer=" & bBuffer & " toLogFile=" & bToLogFile
    Call log(mymsg, LogLEVEL.BASIC)
Exit Sub

Errhandler:
    Debug.Print "Error reading Properties File: " & path & " " & Err.Number & " " & Err.Description
    
End Sub
'delete log file currently set
Public Sub deleteLogFile()
  On Error GoTo Errhandler:
  If (FileExists(Me.LogFilePath)) Then
     Kill (Me.LogFilePath)
  End If
Exit Sub

Errhandler:
  Debug.Print "Error deleting Logfile " & Me.LogFilePath & " " & Err.Number & " " & Err.Description
  
End Sub
'set logfilepath
'- will delete an existing log file if bDelLogFileAtSetup is set to true
Public Sub setLogFile(filePath As String, delExitingFile As Boolean)
  On Error GoTo Errhandler:
  
  Me.LogFilePath = filePath
  'delete if set to true
  If (delExitingFile) Then Call deleteLogFile
  If (bToLogFile) Then Debug.Print "Logfile set to: " & LogFilePath
Exit Sub

Errhandler:
  Debug.Print "Error setLogFile " & LogFilePath & " " & Err.Number & " " & Err.Description
 
End Sub

Public Function getLogParamsFromFile()

  On Error GoTo Errhandler:
  If (FileExists(PropsFileName)) Then
     Debug.Print "Reading: " & PropsFileName
     'read and set parameter from properties file
     readPropertiesFile (PropsFileName)
     getLogParamsFromFile = True
   Exit Function
  End If
  getLogParamsFromFile = False

Exit Function
Errhandler:
  Debug.Print "Error getLogParamsFromFile " & PropsFileName & " " & Err.Number & " " & Err.Description
  getLogParamsFromFile = False
End Function

Private Sub setLogLevel(level As String)
  Dim mylevel
  mylevel = UCase(level)

  Select Case mylevel
    Case "DISABLED":
       iLogLevel = LogLEVEL.DISABLED
    Case "BASIC":
       iLogLevel = LogLEVEL.BASIC
    Case "INFO":
       iLogLevel = LogLEVEL.INFO
    Case "WARN":
       iLogLevel = LogLEVEL.WARN
    Case "FATAL":
       iLogLevel = LogLEVEL.FATAL
    Case "FINE":
       iLogLevel = LogLEVEL.FINE
    Case "FINER":
       iLogLevel = LogLEVEL.FINER
    Case "FINEST":
       iLogLevel = LogLEVEL.FINEST
    Case "ALL":
       iLogLevel = LogLEVEL.ALL
  End Select
End Sub

Public Function getLogBuffer() As String
   getLogBuffer = cLogbuffer.strLogbuffer
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Utils
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'-- extract full path of parent folder of file mypath
Public Function getParentFolder(mypath As String) As String
  Dim pos As Integer
  Dim fullpath As String
  
  pos = InStrRev(mypath, "\")
  If (pos <> 0) Then
    getParentFolder = Left(mypath, pos - 1)
    Exit Function
  End If
  getParentFolder = ""
End Function

'-- Check File Exists --
Public Function FileExists(path As String) As Boolean
  FileExists = (Dir(path) <> "")
End Function

'-- isNothing --
Function checkIsNothing(obj As Object)
  If (obj Is Nothing) Then
     checkIsNothing = True
  Else
     checkIsNothing = False
  End If
End Function

'-- parameters ---
Function readParameter(line As String) As String()
  Dim txtarr() As String
  Dim proparray(2) As String
  
  txtarr = VBA.Split(line, "=")
  If (UBound(txtarr) > 0) Then
    proparray(0) = VBA.Trim(txtarr(0))
    proparray(1) = VBA.Trim(txtarr(1))
    readParameter = proparray
  Else
    readParameter = proparray
  End If
End Function

'-- check text coded boolean value (if read from text file) --
Function valIsTrue(boolval As String) As Boolean
    If ("TRUE" = VBA.UCase(boolval)) Then
        valIsTrue = True
        Exit Function
    End If
    valIsTrue = False
End Function
