VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logbuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' CLASS_MODULE:   VBA Logbuffer - allows a Reference e.g. the 'Set' Method on a String Property
'''
'''                 - see the 'Logging' Module for Usage: the Logging Module
'''                   automatically creates a 'Logger' instance and provides additional
'''                   Features
'''
'''                 - use Macro 'Test' from 'TestLogging' for testing and as an example
'''
''' Date            Developer               Action
''' --------------------------------------------------------------------------
''' 28/08/08        Christian Bolterauer    Created
'''


Public strLogbuffer As String

Private Sub Class_Initialize()
   strLogbuffer = ""
End Sub

Public Sub addline(logmsg As String)
   If (Len(strLogbuffer) > 0) Then
     strLogbuffer = strLogbuffer & vbLf & logmsg
   Else
     strLogbuffer = logmsg  'avoid empty line when strLogbuffer=""
   End If
End Sub

Public Sub writeLogBufferToTraceFile(myfilePath As String)
    
    On Error GoTo Errhandler:
    Dim lines() As String
    Dim line As Variant
    
    If Len(myfilePath) = 0 Then
       Debug.Print "Error: Trace file path is empty."
       Exit Sub
    End If
       
    Dim FileNum As Integer
    FileNum = FreeFile ' next file number
    Open myfilePath For Output As #FileNum ' creates the file if it doesn't exist
    lines = VBA.Split(Me.strLogbuffer, VBA.vbLf)
    For Each line In lines
       Print #FileNum, line ' write Logbuffer to text file
    Next line
    Close #FileNum ' close the file
Exit Sub

Errhandler:
    Debug.Print "Error writing to Tracefile: " & myfilePath & " " & Err.Number & " " & Err.Description
    
End Sub

