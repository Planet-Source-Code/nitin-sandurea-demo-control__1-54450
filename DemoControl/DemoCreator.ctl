VERSION 5.00
Begin VB.UserControl DemoCreator 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H000000FF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   KeyPreview      =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "DemoCreator.ctx":0000
End
Attribute VB_Name = "DemoCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Default Property Values:
Const m_def_ProductName = "DEMO"
Const m_def_Status = "0"
Const m_def_NoOfDays = 0
Const m_def_NofTimes = 0
'Property Variables:
Dim m_ProductName As String
Dim m_Status As StatusData

Dim m_NoOfDays As Integer
Dim m_NofTimes As Integer
Dim DemoExpired As Boolean
Dim InstallDate As String
Dim UsedTimes As Integer
Public Event DemoCompleted(msg As String)

Enum StatusData
  Demo = 0
  Full = 1
End Enum

Dim mobjini As New ClsINI


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get status() As StatusData
    status = m_Status
End Property

Public Property Let status(ByVal New_Status As StatusData)
    m_Status = New_Status
    PropertyChanged "Status"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=3,0,0,0
Public Property Get NoOfDays() As Integer
    NoOfDays = m_NoOfDays
End Property

Public Property Let NoOfDays(ByVal New_NoOfDays As Integer)
    m_NoOfDays = New_NoOfDays
    PropertyChanged "NoOfDays"
        
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get NofTimes() As Integer
    NofTimes = m_NofTimes
End Property

Public Property Let NofTimes(ByVal New_NofTimes As Integer)
    m_NofTimes = New_NofTimes
    PropertyChanged "NofTimes"
End Property



'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Status = m_def_Status
    m_NoOfDays = m_def_NoOfDays
    m_NofTimes = m_def_NofTimes
    m_ProductName = m_def_ProductName
End Sub


'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Status = PropBag.ReadProperty("Status", m_def_Status)
    m_NoOfDays = PropBag.ReadProperty("NoOfDays", m_def_NoOfDays)
    m_NofTimes = PropBag.ReadProperty("NofTimes", m_def_NofTimes)
    m_ProductName = PropBag.ReadProperty("ProductName", m_def_ProductName)
           
           
    If Ambient.UserMode = False Then Exit Sub
    
    
    
    InstallDate = ""
    
    UsedTimes = 0
    
    If m_Status = Demo And m_ProductName <> "" Then
       Dim Key As Long
       Dim PathValue As String
        
        
         InstallDate = mobjini.GetInstallDate(m_ProductName)
         UsedTimes = Val(mobjini.GetUsedTimes(m_ProductName))
         
         
         
         If InstallDate = Empty And UsedTimes = 0 Then
            Call mobjini.WriteInstallDate(m_ProductName, Now)
            Call mobjini.WriteUsedTimes(m_ProductName, 1)
            'Call SaveString(HKEY_LOCAL_MACHINE, "\" & m_ProductName & "\NoOfTimes", "Times", "1")
            InstallDate = Format(Now, "dd/mm/yyyy")
         End If
         
        If UsedTimes < m_NofTimes Then
           UsedTimes = UsedTimes + 1
           Call mobjini.WriteUsedTimes(m_ProductName, Str(UsedTimes))
        Else
        
         DemoExpired = True
          
        End If
       
        If DateDiff("d", InstallDate, Format(Now, "dd/mm/yyyy")) < m_NoOfDays Then
           InstallDate = DateAdd("d", 1, InstallDate)
           'Call SaveString(HKEY_LOCAL_MACHINE, m_ProductName, NoOfTimes, Str(UsedTimes))
        Else
        
        DemoExpired = True
         
        End If
  End If
    
    
End Sub



Private Sub UserControl_Resize()
UserControl.Width = 100
UserControl.Height = 100
End Sub

Private Sub UserControl_Show()
If DemoExpired = True Then
  RaiseEvent DemoCompleted(m_ProductName & " 's demo expired")
End If
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    
    Call PropBag.WriteProperty("Status", m_Status, m_def_Status)
    Call PropBag.WriteProperty("NoOfDays", m_NoOfDays, m_def_NoOfDays)
    Call PropBag.WriteProperty("NofTimes", m_NofTimes, m_def_NofTimes)
    Call PropBag.WriteProperty("ProductName", m_ProductName, m_def_ProductName)
    
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get ProductName() As String
    ProductName = m_ProductName
End Property

Public Property Let ProductName(ByVal New_ProductName As String)
    m_ProductName = New_ProductName
    PropertyChanged "ProductName"
End Property

