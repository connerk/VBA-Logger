Attribute VB_Name = "Logging"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Contents:   Logging Modul for VBA - uses 'LOGGER' Class
'''
''' Comments:   Facade for Logger, with static reference to a Logger instance
'''             The Static Logger allows to write log statments to a logbuffer
'''             that can be read for example inside Errorhandling
'''
''' Example: Replacing Debug.Print:
'''     if you use the Logging Module no initialization needs to be done:
'''     instead 'Debug.Print txt' use: 'Logging.log (txt)'
'''
''' Example: Log4VBA sytle logging:
'''
'''     Dim myLogger As Object 'globaly define
'''
'''     Set myLogger = Logging.getNewLogger(Application.VBE.ActiveVBProject.name) ' initalize Logger and set Module Name for example 'VBALogger'
'''     Call myLogger.setLoggigParams(Logging.lgALL, True, True, True)            ' log  ALL to Console, Buffer, File
'''
'''     myLogger.logINFO "This is my message ..", "MySubOrFunction"               ' log a message in Sub 'MySubOrFunction'
'''
'''     Result:
'''     (28.08.2008 10:53:20)[VBALogger::MySubOrFunction]-INFO:  This is my message ..
'''
''' Changing Settings:
'''
''' The bestway to change Loglevels and the settings logging to console, buffer, or logfile
''' is by changing the settings via properties file "vba_log.properties"
''' With this version the properties file is expected in the same directory as the Module
''' containing the LOGGER Class (and the Logging Module)
''' Example:
''' ---------------------------------------------
''' #
''' # -- settings for VBA logging --
''' #
''' # LOG_LEVEL:
''' #
''' #  DISABLED
''' #  BASIC 'like Debug.Print
''' #  FATAL
''' #  WARN
''' #  INFO
''' #  FINE
''' #  FINER
''' #  FINEST
''' #  ALL
''' #
''' LOG_LEVEL = info
''' LOG_TO_CONSOLE = True
''' LOG_TO_BUFFER = True
''' LOG_TO_FILE = True
''' #  Default LOG_FILE_PATH is the same place as Project File containing the Logger Modul
''' #LOG_FILE_PATH=C:\vba_logger.log
''' -----------------------------------------
'''
''' Settings can be changed using vba code with the setLoggigParams(..) procedure
''' example:
'''       Call Logging.setLoggigParams(Logging.lgBASIC, True, True, False)
'''
''' Example use for LogBuffer:
'''       If (Err) Then Logging.writeLogBufferToTraceFile
'''
'''
''' Date        Developer               Action
''' --------------------------------------------------------------------------
''' 28/08/08    Christian Bolterauer    Created
'''
'''

Option Explicit

' global to allow access to Logger Class instance via Logging Module
Public defaultLogger As Logger

'copy of levels from Logger Class to expose levels via the Logging Module
'Note that the enum 'LogLEVEL' is only visable within the VBAProject that contains the Logger Class.
'The Const variables are visable to every Modul where Logging can be accessed
Public Const lgDISABLED = LogLEVEL.DISABLED
Public Const lgBASIC = LogLEVEL.BASIC
Public Const lgFATAL = LogLEVEL.FATAL
Public Const lgWARN = LogLEVEL.WARN
Public Const lgINFO = LogLEVEL.INFO
Public Const lgFINE = LogLEVEL.FINE
Public Const lgFINER = LogLEVEL.FINER
Public Const lgFINEST = LogLEVEL.FINEST
Public Const lgALL = LogLEVEL.ALL

'setter for prime logparameters
Sub setLoggigParams(myloglevel As Integer, toConsole As Boolean, toBuffer As Boolean, toLogFile As Boolean)
  If (myloglevel = LogLEVEL.DISABLED) Then Debug.Print "Logging is disabled."
  
  'Important: initilaze logger by calling log() before setting params
  log ("Logging with logLevel=" & defaultLogger.getLogLevelName(myloglevel) & " ToConsole=" & toConsole & " ToBuffer=" & toBuffer & " ToLogFile=" & toLogFile)
  Call defaultLogger.setLoggigParams(myloglevel, toConsole, toBuffer, toLogFile)
  
  'Inital LogfilePath set here
  'Call defaultLogger.setLogFile(Application.ActiveWorkbook.path & "\vba_logger.log")

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Static defaultLogger instance
'
' The live time of this logger instance is as long as the application runs
' This allows to write log messages to a buffer that can be processed even if modules are changed
'
' The defaultLogger is initialized the first time when any of the following log statements is called
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub thislog(msg As String, myloglevel As LogLEVEL, Optional slogpoint As String)
  Static mydefaultLogger As New Logger 'singelton
  
  '- if static value is not set assume start of vba session and delete the log file -
  If (defaultLogger Is Nothing) Then
     Call mydefaultLogger.deleteLogFile
  End If
  
  Call mydefaultLogger.log(msg, myloglevel, slogpoint)
  Set defaultLogger = mydefaultLogger  'refence to static object
End Sub

Public Sub log(sLogText As String, Optional slogpoint As String)
   Call thislog(sLogText, LogLEVEL.BASIC, slogpoint)
End Sub
Public Sub logINFO(sLogText As String, Optional slogpoint As String)
   Call thislog(sLogText, LogLEVEL.INFO, slogpoint)
End Sub
Public Sub logWARN(sLogText As String, Optional slogpoint As String)
   Call thislog(sLogText, LogLEVEL.WARN, slogpoint)
End Sub
Public Sub logFATAL(sLogText As String, Optional slogpoint As String)
   Call thislog(sLogText, LogLEVEL.FATAL, slogpoint)
End Sub
Public Sub logFINE(sLogText As String, Optional slogpoint As String)
   Call thislog(sLogText, LogLEVEL.FINE, slogpoint)
End Sub
Public Sub logFINER(sLogText As String, Optional slogpoint As String)
   Call thislog(sLogText, LogLEVEL.FINER, slogpoint)
End Sub
Public Sub logFINEST(sLogText As String, Optional slogpoint As String)
   Call thislog(sLogText, LogLEVEL.FINEST, slogpoint)
End Sub
Function getLogBuffer()
   If (defaultLogger Is Nothing) Then
      'initilize defaultLogger calling ..
      Call thislog("Retrieving LogBuffer..", LogLEVEL.FINE)
   End If
   getLogBuffer = defaultLogger.getLogBuffer
End Function
'set setModulName: ensures that defaultLogger is initalized before value is set
Public Sub setModulName(myModulName As String)
    If (defaultLogger Is Nothing) Then
      'initilize defaultLogger calling ..
      Call thislog("Setting ModulName to " & myModulName, LogLEVEL.FINE)
    End If
    defaultLogger.ModulName = myModulName
End Sub
Public Sub writeLogBufferToTraceFile(Optional myfilePath As String)
    If (defaultLogger Is Nothing) Then
      'initilize defaultLogger calling ..
      Call thislog("Writing LogBuffer to TraceFile ..", LogLEVEL.FINE)
   End If
   defaultLogger.writeLogBufferToTraceFile (myfilePath)
End Sub

'*******************************************************************************************
'* MODULE:    getNewLogger
'*
'* PURPOSE:   Return a logger object with the defaults set.
'*            The Log Buffer of the new Logger created by this factory method is set
'*            to defaultLogger.strLogbuffer so that all log entries of a session can be traced
'*
'* PARAMETERS: sModulName - the VBA Module that will be used as an identifier within the log file.
'*******************************************************************************************
Public Static Function getNewLogger(sModulName As String) As Logger
    Dim myLogger As New Logger
    myLogger.ModulName = sModulName
    
    'set the logBuffer to defaultLogger Logbuffer so that all log entries of a session can be traced
    Set myLogger.cLogbuffer = defaultLogger.cLogbuffer
    
    Set getNewLogger = myLogger
End Function





