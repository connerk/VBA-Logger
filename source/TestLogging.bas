Attribute VB_Name = "TestLogging"
'''
''' Basic test macro to test Logger Class using Logging
'''

'define 'myLogger' as 'Object' (not 'Logger') to ensure that
'this test Class works in VBAProjects that reference 'Logging.xla'
'since Public Class Moduls may not be exposed as Type between VBAProjects
Dim myLogger As Object


Sub Test()
  
 Logging.setModulName (Application.VBE.ActiveVBProject.name)
 Logging.logINFO ("***Starting Logger test..")
 
 Call printLogLevels

 Logging.log ("***Testing LogLevels..")
 Call Logging.setLoggigParams(Logging.lgALL, True, True, True)
 Call printLogLevels
 Call Logging.setLoggigParams(Logging.lgFINEST, True, True, True)
 Call printLogLevels
 Call Logging.setLoggigParams(Logging.lgFINER, True, True, True)
 Call printLogLevels
 Call Logging.setLoggigParams(Logging.lgFINE, True, True, True)
 Call printLogLevels
 Call Logging.setLoggigParams(Logging.lgINFO, True, True, True)
 Call printLogLevels
 Call Logging.setLoggigParams(Logging.lgWARN, True, True, True)
 Call printLogLevels
 Call Logging.setLoggigParams(Logging.lgFATAL, True, True, True)
 Call printLogLevels
 Call Logging.setLoggigParams(Logging.lgBASIC, True, True, True)
 Call printLogLevels
 Logging.log ("***Now Turn logging off ..")
 Call Logging.setLoggigParams(Logging.lgDISABLED, True, True, True)
 Call printLogLevels
 Call Logging.setLoggigParams(Logging.lgALL, True, True, True)
 Call Logging.log("***Testing logging with 'logpoint' entry ..")
 Call printLogLevelsWithLogPoint
 
 Call Logging.setLoggigParams(Logging.lgALL, True, False, False)
 Logging.log ("***Testing logBuffer ..")
 Logging.log "----Printing Logging.getLogBuffer to Console only ----"
 Logging.log Logging.getLogBuffer
 Call Logging.setLoggigParams(Logging.lgALL, True, True, True)
 
 Logging.setModulName ("")
 
 Call TestLoggerInstance
 
 Logging.log ("***Testing writing logBuffer to Tracefile ..")
 Logging.writeLogBufferToTraceFile
 
 Logging.log ("***Testing done.***")
 
End Sub


Private Sub printLogLevels()
 Logging.log ("-LogBasic = like Debug.Print-")
 Logging.logINFO ("-logINFO-")
 Logging.logWARN ("-logWARN-")
 Logging.logFATAL ("-logFATAL-")
 Logging.logFINE ("-logFINE-")
 Logging.logFINER ("-logFINER-")
 Logging.logFINEST ("-logFINEST-")
End Sub

Private Sub printLogLevelsWithLogPoint()
 Logging.log ("-LogBasic = like Debug.Print-")
 Logging.logINFO "-logINFO-", "printLogLevelsWithLogPoint"
 Logging.logWARN "-logWARN-", "printLogLevelsWithLogPoint"
 Logging.logFATAL "-logFATAL-", "printLogLevelsWithLogPoint"
 Logging.logFINE "-logFINE-", "printLogLevelsWithLogPoint"
 Logging.logFINER "-logFINER-", "printLogLevelsWithLogPoint"
 Logging.logFINEST "-logFINEST-", "printLogLevelsWithLogPoint"
End Sub

Sub TestLoggerInstance()
  
  Set myLogger = Logging.getNewLogger(Application.VBE.ActiveVBProject.name)
  Call myLogger.setLoggigParams(Logging.lgALL, True, True, True)
  myLogger.logBASIC "***Starting TestLoggerInstance test.."
  myLogger.logBASIC "-LogBasic = like Debug.Print-", "TestLoggerInstance"
  myLogger.logINFO "-logINFO-", "TestLoggerInstance"
  myLogger.logWARN "-logWARN-", "TestLoggerInstance"
  myLogger.logFATAL "-logFATAL-", "TestLoggerInstance"
  myLogger.logFINE "-logFINE-", "TestLoggerInstance"
  myLogger.logFINER "-logFINER-", "TestLoggerInstance"
  myLogger.logFINEST "-logFINEST-", "TestLoggerInstance"
  
  'call a sub
  Call MySubOrFunction
  
  Call myLogger.setLoggigParams(Logging.lgALL, True, False, False)
  myLogger.logBASIC "*** printing the TestLoggerInstance buffer to Console.."
  myLogger.logBASIC myLogger.getLogBuffer
  
End Sub

Sub MySubOrFunction()
  myLogger.logINFO "This is my message ..", "MySubOrFunction"               ' log a message in Sub 'MySubOrFunction'
End Sub

