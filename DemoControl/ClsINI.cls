VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Public Function SetModeToRemote()
    On Error GoTo ErrorHandler
    
        Dim iniPath As String
    
        iniPath = App.Path + "\SHUTBACKLOGDATA.DLL"
        
        Call WriteValueToIni(iniPath, "SETTINGS", "PREVCONNECTION", "REMOTE")
        
        Exit Function
        
ErrorHandler:

    MsgBox Err.Description & Err.Number
        
End Function
Public Function GetInstallDate(ByVal ProductName As String) As String

On Error GoTo ErrorHandler
    
        Dim iniPath As String
    
        iniPath = App.Path + "\SHUTBACKLOGDATA.DLL"
        
        GetInstallDate = GetValueFromInI(iniPath, ProductName, "INSTALLDATE")
        
        Exit Function
        
ErrorHandler:

    MsgBox Err.Description & Err.Number
        
    
End Function
Public Function GetNoOfTimes(ByVal ProductName As String) As String

On Error GoTo ErrorHandler
    
        Dim iniPath As String
    
        iniPath = App.Path + "\SHUTBACKLOGDATA.DLL"
        
        GetNoOfTimes = GetValueFromInI(iniPath, ProductName, "NOOFTIMES")
        
        Exit Function
        
ErrorHandler:

    MsgBox Err.Description & Err.Number
        
    
End Function
Public Function GetNoOfDAYS(ByVal ProductName As String) As String

On Error GoTo ErrorHandler
    
        Dim iniPath As String
    
        iniPath = App.Path + "\SHUTBACKLOGDATA.DLL"
        
        GetNoOfDAYS = GetValueFromInI(iniPath, ProductName, "NOOFDAYS")
        
        Exit Function
        
ErrorHandler:

    MsgBox Err.Description & Err.Number
        
    
End Function
Public Function GetUsedTimes(ByVal ProductName As String) As String

On Error GoTo ErrorHandler
    
        Dim iniPath As String
    
        iniPath = App.Path + "\SHUTBACKLOGDATA.DLL"
        
        GetUsedTimes = GetValueFromInI(iniPath, ProductName, "USEDTIMES")
        
        Exit Function
        
ErrorHandler:

    MsgBox Err.Description & Err.Number
        
    
End Function

Public Function WriteInstallDate(ByVal ProductName As String, ByVal Value As String)

    On Error GoTo ErrorHandler
    
        Dim iniPath As String
    
        iniPath = App.Path + "\SHUTBACKLOGDATA.DLL"
        
        Call WriteValueToIni(iniPath, ProductName, "INSTALLDATE", Format(Value, "dd/mm/yyyy"))
        
        Exit Function
        
ErrorHandler:

    MsgBox Err.Description & Err.Number
End Function
Public Function WriteUsedTimes(ByVal ProductName As String, ByVal Value As String)

    On Error GoTo ErrorHandler
    
        Dim iniPath As String
    
        iniPath = App.Path + "\SHUTBACKLOGDATA.DLL"
        
        Call WriteValueToIni(iniPath, ProductName, "USEDTIMES", Value)
        
        Exit Function
        
ErrorHandler:

    MsgBox Err.Description & Err.Number
End Function

Private Sub Class_Initialize()

 Dim FileNum As String

  FileNum = FreeFile
  Open App.Path & "\SHUTBACKLOGDATA.DLL" For Append As FileNum
  Close #FileNum
 
 End Sub
 
