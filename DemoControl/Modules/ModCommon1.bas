Attribute VB_Name = "ModCommon"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3C4B298F02D0"

'Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Public Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020




Public Const GWL_ID = (-12)

Public Const GWL_STYLE = (-16)

Public Const WS_DLGFRAME = &H400000

Public Const WS_SYSMENU = &H80000

Public Const WS_MINIMIZEBOX = &H20000

Public Const WS_MAXIMIZEBOX = &H10000

Public Enum GRADIENTS
  GRAD_LEFTTORIGHT = 0
  GRAD_TOPTOBOTTOM = 1
End Enum

Const MAX_OPTION_NAME = 128


Private Const SWP_HIDEWINDOW = &H80

#If Win32 Then
      Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
      #Else
         Public Declare Function SendMessage Lib "User" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wp As Integer, lp As Any) As Long
      #End If

      

Public Const LB_SELECTSTRING = &H18C
Public Const LB_FINDSTRING = &H18F
Public Const LB_SETCURSEL = &H186
'##ModelId=3C4B298F030D
Private Const SWP_SHOWWINDOW = &H40

Public Const CB_SELECTSTRING = &H14D
Public Const CB_SETCURSEL = &H14E
Public Const CB_FINDSTRING = &H14C
'Public Const CB_SHOWDROPDOWN = &H14F
Const WM_USER = &H400
Global Const CB_SHOWDROPDOWN = WM_USER + 15
Public Declare Function timeSetEvent Lib "winmm.dll" (ByVal uDelay As Long, ByVal uResolution As Long, ByVal lpFunction As Long, ByVal dwUser As Long, ByVal uFlags As Long) As Long


Enum MESSAGES
   NO = 0
   YES = 1
   ABORT = 2
   RETRY = 3
   OK = 4
End Enum
  
Public FormMessage As MESSAGES
Public Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer




Public Sub MakeTranslucent(Who As Form, Optional tColor As Long) 'Was (Who as Object) before...

On Local Error Resume Next

Dim HW As Long
Dim HA As Long
Dim iLeft As Integer
Dim iTop As Integer
Dim iWidth As Integer
Dim iHeight As Integer

If IsMissing(tColor) Or tColor = 0 Then
    tColor = RGB(0, 0, 200)
End If

Who.AutoRedraw = True
Who.Hide

DoEvents

HW = GetDesktopWindow()
HA = GetDC(HW)

'Get the Left, Top, Width and Height of the Form...
iLeft = Who.Left / Screen.TwipsPerPixelX
iTop = Who.Top / Screen.TwipsPerPixelY '+ 25    If using a form with a titlebar (border)...
iWidth = Who.ScaleWidth
iHeight = Who.ScaleHeight

'Now, Transfer the contents of the Desktop Window to the Form...
Call BitBlt(Who.hDC, 0, 0, iWidth, iHeight, HA, iLeft, iTop, SRCCOPY) 'iLeft + 4    If using a form with a titlebar (border)...

'Show...
Who.Picture = Who.Image
Who.Show

'Release the DC...
Call ReleaseDC(HW, HA)

'Add color...
Who.DrawMode = 9
Who.ForeColor = tColor
Who.Line (0, 0)-(iWidth, iHeight), , BF

End Sub

'##ModelId=3C4B298F0381
Public Sub SetComboItem(cmbo As Control, Data As String)

Dim i As Integer

    For i = 0 To cmbo.ListCount - 1
        Debug.Print cmbo.List(i)
        If cmbo.List(i) = Data Then
            cmbo.ListIndex = i
           Exit Sub
        End If
    Next
      cmbo.ListIndex = -1
      
End Sub




'##ModelId=3C4B298F03AC
Public Function IsValidEmail(TxtEmail As Control, Optional ByRef msg As String) As Boolean

        Dim i As String
        Dim j As Integer
        Dim Length As Integer
        Dim m As String
        Dim POS As Integer
        Dim length1 As Integer
    
       If msg <> Empty Then
        
            Length = Len(TxtEmail.Text)
          
          'IsValidEmail the value of email fields
            
            If TxtEmail.Text <> "" Then
            
                i = TxtEmail.Text
                j = InStr(1, i, "@", vbTextCompare)
                
            End If
            
            If j = 1 Then
            
                msg = " @ Can't be first character"
                IsValidEmail = False
                Exit Function
                
            End If
            
            If j = Length Then
                msg = "@ Can't be last character"
                IsValidEmail = False
                Exit Function
            End If
        
            If j = 0 Then
            
                msg = "Invalid E mail Address"
                IsValidEmail = False
                Exit Function
                
            End If
        
        
            'IsValidEmailing the . in email field
            
            m = TxtEmail.Text
               
            length1 = Len(m)
               
            POS = InStr(1, m, ".")
            
            If POS = 0 Then
            
                msg = "The email address is incorrect"
                IsValidEmail = False
                Exit Function
             
             ElseIf POS = length1 Then
             
                    msg = ". Can't be last character"
                    IsValidEmail = False
                    Exit Function
             
             ElseIf POS = 1 Then
             
                    msg = ".Can't be first character"
                    IsValidEmail = False
                    Exit Function
                    
             End If
             
        Else
          
            Length = Len(TxtEmail.Text)
            'IsValidEmail the value of email fields
        If TxtEmail.Text <> "" Then
            i = TxtEmail.Text
            j = InStr(1, i, "@", vbTextCompare)
        End If
        
        If j = 1 Then
            IsValidEmail = False
            Exit Function
        End If
         If j = Length Then
            IsValidEmail = False
            Exit Function
        End If
        
        If j = 0 Then
            IsValidEmail = False
            Exit Function
                
        End If
        
        
            'IsValidEmailing the . in email field
            m = TxtEmail.Text
            length1 = Len(m)
            POS = InStr(1, m, ".")
            
             If POS = 0 Then
                 IsValidEmail = False
                 Exit Function
                 
             ElseIf POS = length1 Then
             
                IsValidEmail = False
                Exit Function
                
             ElseIf POS = 1 Then
             
                IsValidEmail = False
                Exit Function
                
             End If
            
             IsValidEmail = True
        
        End If
    End Function
'##ModelId=3C4B298F03AF
Public Sub CenterScreen(aForm As Form, Optional MFrm As Object)
If Not MFrm Is Nothing Then
  aForm.Move (MFrm.Width - aForm.Width) / 2, (MFrm.Height - aForm.Height) / 2
Else
 aForm.Move (Screen.Width - aForm.Width) / 2, (Screen.Height - aForm.Height) / 2
 End If
End Sub
      '##ModelId=3C4B298F03B1
      Function OfficeClosed(TheDate) As Integer
      OfficeClosed = False
    ' Test for Saturday or Sunday.
         
         If Weekday(TheDate) = 1 Or Weekday(TheDate) = 7 Then
            OfficeClosed = True
            ' Test for Holiday.
         End If

      End Function


Public Sub HideTaskbar()

Dim Rs As Long

Rs = FindWindow("Shell_traywnd", "")
Call SetWindowPos(Rs, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)

End Sub


Public Sub ShowTaskbar()

Dim Rs As Long

Rs = FindWindow("Shell_traywnd", "")
Call SetWindowPos(Rs, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)

End Sub


Public Sub CheckDate()


Dim Flreg As Boolean
Dim myfile As String
Dim ft As String
Dim f As String * 255
Dim G As String
Dim D As Long
    ft = Environ("windir")
    ft = ft & "\win.ini"
'sShortDate
'd = GetPrivateProfileString("intl", "sShortDate", "f", "f", 255, ft)
    D = GetPrivateProfileString("intl", "sShortDate", G, f, 255, ft)

'MsgBox f
    If StrComp(f, "DD/MM/YYYY", vbTextCompare) <> 0 Then
        Call MsgBox("Convert The System Date Format To [DD/MM/YYYY]", vbApplicationModal + vbInformation, "Payroll Management")
      
    End If

    Exit Sub
End Sub



Public Function LastDateOfMonth(D As String) As String
         
    LastDateOfMonth = DateSerial(Year(D), Month(D) + 1, 0)
    
End Function

Public Function IsComboItem(cmbo As Control, Data As String) As Boolean
    Dim i As Integer

    For i = 0 To cmbo.ListCount - 1

        If cmbo.List(i) = Data Then
            IsComboItem = True
           Exit Function
        End If
    Next
      IsComboItem = False
End Function

Public Function DaysInMonth(ByVal D As Date) As Long
      ' Requires a date argument because February can change
      ' if it's a leap year.
        Select Case Month(D)
          Case 2
            If LeapYear(Year(D)) Then
              DaysInMonth = 29
            Else
              DaysInMonth = 28
            End If
          Case 4, 6, 9, 11
            DaysInMonth = 30
          Case 1, 3, 5, 7, 8, 10, 12
            DaysInMonth = 31
        End Select
      End Function


Public Function LeapYear(ByVal YYYY As Long) As Boolean
        LeapYear = YYYY Mod 4 = 0 _
                   And (YYYY Mod 100 <> 0 Or YYYY Mod 400 = 0)
End Function


Public Function RETURNMONTHNO(NameMonth) As Integer

    Dim MONTHNO As Integer
            
            
    NameMonth = UCase(NameMonth)
        
        
        
        Select Case (NameMonth)

        Case "JANUARY"

            MONTHNO = 1

        Case "FEBRUARY"

            MONTHNO = 2


        Case "MARCH"

            MONTHNO = 3

        Case "APRIL"

            MONTHNO = 4

        Case "MAY"

            MONTHNO = 5

        Case "JUNE"

            MONTHNO = 6

        Case "JULY"

            MONTHNO = 7

        Case "AUGUST"

            MONTHNO = 8

        Case "SEPTEMBER"

            MONTHNO = 9

        Case "OCTOBER"

            MONTHNO = 10

        Case "NOVEMBER"

            MONTHNO = 11

        Case "DECEMBER"

            MONTHNO = 12

    End Select

    RETURNMONTHNO = MONTHNO

End Function


Public Function AddMonthsToCombo(cmbo As Control) As Boolean


On Error GoTo ErrorHandler


   cmbo.Clear
   Dim i As Integer
   For i = 1 To 12
       cmbo.AddItem MonthName(i)
   Next
   
   Exit Function
       
       
ErrorHandler:
    MsgBox Err.Description & Err.Number

End Function


Public Function NoOfChar(Data As String, FindData As String) As String

    Dim Length As Integer
    Dim i As Integer
    Dim Found As Integer
    
     
     Found = 0
     
     Length = Len(Data)
     
    
     
    For i = 1 To Length
            
        If Mid$(Data, i, 1) = FindData Then
           Found = Found + 1
        End If
        
    Next
    
     
    NoOfChar = Found
    
    
End Function


Public Function Even(Data As Integer) As Boolean

    If Data Mod 2 = 0 Then
    
        Even = True
        Exit Function
        
    End If
    
    Even = False
    
End Function


Public Function Encrypt(Data As String, Key As Integer) As String

Dim TData As String
Dim i As Integer
For i = 1 To Len(Data)
    
    TData = TData + Chr(Asc(Mid$(Data, i, i)) Xor Key)
Next

Encrypt = TData

End Function


Public Function ListItem(hwnd As Long, Data As String) As Long

'Data = Data + Chr(0)

ListItem = SendMessage(hwnd, LB_SELECTSTRING, -1, CStr(Data))

End Function

Public Function CmboItem(hwnd As Long, Data As String) As Long

'Data = Data + Chr(0)

CmboItem = SendMessage(hwnd, CB_SELECTSTRING, -1, CStr(Data))

End Function
Public Function CmboDrop(hwnd As Long, status As Boolean) As Long

'Data = Data + Chr(0)

CmboDrop = SendMessage(hwnd, CB_SHOWDROPDOWN, -1, 0&)

End Function


Public Function GetValueFromInI(Filename As String, Section As String, Ikey As String)
    
    Dim lresult As Long
    Dim Rdata As String
    Rdata = String(128, Chr(0))
    Section = Section + Chr(0)
    Ikey = Ikey + Chr(0)
    Filename = Filename + Chr(0)
    lresult = GetPrivateProfileString(Section, Ikey, " ", Rdata, 128, Filename)
    
    GetValueFromInI = Left(Rdata, lresult)
End Function

Public Function WriteValueToIni(Filename As String, Section As String, Ikey As String, Data As String)

    Dim lresult As Long
    
    Section = Section + Chr(0)
    Ikey = Ikey + Chr(0)
    Data = Data + Chr(0)
    lresult = WritePrivateProfileString(Section, Ikey, Data, Filename)
    
End Function
Public Sub ShowTitleBar(frmIn As Form, bShow As Integer)

     Static iOldMenu As Integer

     Static lSavedStyle As Long

     Dim lNewStyle As Long

     Dim R As Long

 

     If bShow Then

           'get the current style attributes

           lNewStyle = GetWindowLong&(frmIn.hwnd, GWL_STYLE)

 

           'set only the attributes that were removed earlier

           lNewStyle = lNewStyle Or lSavedStyle

 

           're-establish the menu

           If iOldMenu <> 0 Then

                 R = SetWindowWord&(frmIn.hwnd, GWL_ID, iOldMenu)

           End If

 

           'set the new style

           R = SetWindowLong&(frmIn.hwnd, GWL_STYLE, lNewStyle)

 

           'force Visual Basic to update the Form

           frmIn.Move frmIn.Left

           frmIn.Refresh

     Else

           'get the current style attributes

           lNewStyle = GetWindowLong&(frmIn.hwnd, GWL_STYLE)

 

           'determine whether the Form has a dialog frame,

           'a ControlBox, a minimize button, or a maximize

           'button and save this info for later use

           lSavedStyle = 0

           lSavedStyle = lSavedStyle Or (lNewStyle And WS_DLGFRAME)

           lSavedStyle = lSavedStyle Or (lNewStyle And WS_SYSMENU)

           lSavedStyle = lSavedStyle Or (lNewStyle And WS_MINIMIZEBOX)

           lSavedStyle = lSavedStyle Or (lNewStyle And WS_MAXIMIZEBOX)

 

           'remove the attributes for a dialog frame, a

           'ControlBox, a minimize button and a maximize button

           lNewStyle = lNewStyle And Not WS_DLGFRAME

           lNewStyle = lNewStyle And Not WS_SYSMENU

           lNewStyle = lNewStyle And Not WS_MINIMIZEBOX

           lNewStyle = lNewStyle And Not WS_MAXIMIZEBOX
           'is there a menu associated with this Form?
           iOldMenu = GetWindowWord%(frmIn.hwnd, GWW_ID)
           If iOldMenu <> 0 Then
                 'yes-zero it the menu handle
                 R = SetWindowWord&(frmIn.hwnd, GWW_ID, 0)
           End If
           'set the new style
           R = SetWindowLong&(frmIn.hwnd, GWL_STYLE, lNewStyle)
           'force Visual Basic to update the Form and get rid of the
           'title bar
           frmIn.Move frmIn.Left

'           frmIn.Refresh

     End If

End Sub

Sub FadeFormBlue(vForm As Form)
'Example:
'Private Sub Form_Paint()
'FadeFormBlue Me
'End Sub
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
       ' vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0, 255 - intLoop), B
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0, 255 - intLoop), B
        
    Next intLoop
End Sub




Public Sub DrawGradient(Picture As PictureBox, ByVal Orientation As GRADIENTS, ByVal R As Integer, ByVal G As Integer, ByVal B As Integer)
  Dim Count      As Byte
  Dim CycleCount As Integer
  Dim X          As Integer
  
  If Orientation = GRAD_TOPTOBOTTOM Then
    CycleCount = Picture.ScaleHeight \ 100
  Else
    CycleCount = Picture.ScaleWidth \ 100
  End If
  
  For Count = 1 To 100
    X = X + 1
    
    Select Case Orientation
      Case GRAD_TOPTOBOTTOM:
        If X > Picture.ScaleHeight Then Exit For
        Picture.Line (0, X)-(Picture.Width, X + CycleCount - 1), RGB(R, G, B), BF
      Case GRAD_LEFTTORIGHT:
        If X > Picture.ScaleWidth Then Exit For
        Picture.Line (X, 0)-(X + CycleCount - 1, Picture.Height), RGB(R, G, B), BF
        
    End Select
    
    X = X + CycleCount
    
    R = R + 1: If R = 90 Then R = 90
    G = G + 1: If G = 126 Then G = 126
    B = B + 1: If B = 220 Then B = 220
    
  Next Count
End Sub


Public Function AddColorToCombo(cmbo As Control) As Boolean


On Error GoTo ErrorHandler


   cmbo.Clear
   
   Dim i As Integer
   
       cmbo.AddItem "RED"
       cmbo.AddItem "BLUE"
       cmbo.AddItem "GREEN"
       cmbo.AddItem "WHITE"
       cmbo.AddItem "YELLOW"
       cmbo.AddItem "CYAN"
   
   Exit Function
       
       
ErrorHandler:
    MsgBox Err.Description & Err.Number

End Function
Public Function GetColorValue(ColorName As String) As Integer


On Error GoTo ErrorHandler


 
   Dim i As Integer
   
   Select Case ColorName
   
   Case "RED"
        GetColorValue = 4
        Exit Function
    Case "BLUE"
        GetColorValue = 9
        Exit Function
     Case "GREEN"
        GetColorValue = 10
        Exit Function
     Case "YELLOW"
        GetColorValue = 14
        Exit Function
     Case "WHITE"
         GetColorValue = 15
         Exit Function
     Case "CYAN"
         GetColorValue = 11
         Exit Function
    End Select
   
   Exit Function
       
       
ErrorHandler:
    MsgBox Err.Description & Err.Number

End Function

Public Function GetColorName(ColorValue As Long) As String


On Error GoTo ErrorHandler



   Dim i As Integer
   
   Select Case ColorValue
   
     Case 4
          GetColorName = "RED"
          Exit Function
    Case 9
        GetColorName = "BLUE"
        Exit Function
     Case 10
        GetColorName = "GREEN"
        Exit Function
     Case 14
        GetColorName = "YELLOW"
        Exit Function
     Case 15
        GetColorName = "WHITE"
         Exit Function
     Case 11
         GetColorName = "CYAN"
         Exit Function
    End Select
   
   Exit Function
       
       
ErrorHandler:
    MsgBox Err.Description & Err.Number

End Function

Function readBinFile(ByVal bfilename As String)
          Dim fl As Long
          Dim binbyte() As Byte
          Dim binfilestr As String

          On Error GoTo errHandler

          Open bfilename For Binary Access Read As #1
            
            fl = FileLen(bfilename)
            ReDim binbyte(fl)
            Get #1, , binbyte
            Close #1
          
          
          readBinFile = binbyte
          Exit Function

errHandler:
          Exit Function
      End Function

Function WriteBinFile(Filename As String, Data() As Byte)

          Dim fl As Long
          Dim binbyte() As Byte
          Dim binfilestr As String

          On Error GoTo errHandler

          Open Filename For Binary Access Write As #1
            
            
            
               Put #1, , Data
            
               Close #1
                                
          Exit Function

errHandler:
          Exit Function
End Function


