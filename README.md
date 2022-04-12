# TemplateVBA
AOPTemplate内使用的VBA
Option Explicit
    Dim strCalcScript As String
    Dim intCalcsRunSuccessfully As Integer


Sub ConnectedMultipleTabs(i As Integer)
Dim lngConnected, a, b As Long
Dim strUser As String
Dim strWord As String
Dim strDatabase As String
Dim strAppName As String
Dim strServer As String

Dim TempUserEnc As String
Dim TempPassEnc As String
Dim DbEsbUserEnc As String
Dim DbEsbPwdEnc As String

strUser = Sheets("Control").Range("EssUser")
strWord = Sheets("Control").Range("EssWord")
strServer = Sheets("Settings").Range("G" & i)
strAppName = Sheets("Settings").Range("I" & i)
strDatabase = Sheets("Settings").Range("K" & i)

'check server name
If strServer = "" Then
    MsgBox "No Server name for item number:  " & i - 6 & ". Please complete this input to continue.", vbExclamation, DbName
    Sheets("Settings").Select
    End
End If

'check app name
If strAppName = "" Then
    MsgBox "No Application name for item number:  " & i - 6 & ". Please complete this input to continue.", vbExclamation, DbName
    Sheets("Settings").Select
    End
End If

'check db name
If strDatabase = "" Then
    MsgBox "No Database name for item number:  " & i - 6 & ". Please complete this input to continue.", vbExclamation, DbName
    Sheets("Settings").Select
    End
End If

'enable essbase error message display
a = EssVSetGlobalOption(5, 3)

If ActiveSheet.name = "Control" Then
    Exit Sub
Else
    application.DisplayAlerts = False
    
    '*******
    If Sheets("Control").Range("EssUser").Value = "" And Sheets("Control").Range("LoginDataStatus").Value = "Removed" Then
                
        'decrypt login info and login
        TempPassEnc = TempPass
        EncryptDecrypt TempPass

        'read the login info from a temp variable and clean the switch
        DbEsbUser = TempUser
        DbEsbPwd = TempPass
        'store encrypted login info on the workbook
        Sheets("Control").Range("EssUser").Value = TempUser
        Sheets("Control").Range("EssWord").Value = TempPassEnc
        Range("LoginDataStatus").ClearContents
    Else
        'reads login info & decrypt
        TempPass = strWord
        EncryptDecrypt TempPass
        DbEsbPwd = TempPass
        DbEsbUser = strUser
        Range("LoginDataStatus").ClearContents
    End If
        
    '*******
    
    lngConnected = EssVConnect(Empty, DbEsbUser, DbEsbPwd, strServer, strAppName, strDatabase)
    If lngConnected = 0 Then
        Range("OnLineStatus").Value = "On-Line"
        application.StatusBar = "Processing tab " & Sheets("Settings").Range("C" & i).Value
        'application.StatusBar = "Sheet " & ActiveSheet.Name & " Connected to: " & strServer & " " & strAppName & " " & strDatabase
    Else
        MsgBox "Essbase was unbale to connect the following sheet: " & ActiveSheet.name & " using this username: " & strUser & Chr(13) & Chr(13) & " This tab is pointing to the following Server / App: " & strServer & " " & strAppName & " " & strDatabase & Chr(13) & Chr(13) & "Please ensure this username have access to the Server/App/Db. and all server, app, and database are correctly spelled.", vbExclamation, DbName
        
        MsgBox "The Refresh tabs process will be incomplete. For security reasons the Retrieval process will be stoped until this issue is resolved.", vbInformation, DbName
        application.StatusBar = False
        End
        
    End If
    application.DisplayAlerts = True
End If

application.DisplayAlerts = True

End Sub

Sub SetInputSheetOptions()
Dim X As Integer
'Set indent setting to no indentation
    X = EssVSetSheetOption(Null, 5, 1)
    If X <> 0 Then
        MsgBox "Error setting indentation."
    End If
    
'Set adjust columns to off
    X = EssVSetSheetOption(Null, 12, False)
    If X <> 0 Then
        MsgBox "Error setting column adjustment."
    End If
    
'Set alias name usage
    X = EssVSetSheetOption(Null, 13, True)
    If X <> 0 Then
        MsgBox "Error setting alias usage."
    End If
    
'Set styles usage to off
    X = EssVSetSheetOption(Null, 18, False)
    If X <> 0 Then
        MsgBox "Error turning styles off."
    End If
    
'Set #missing label
    If Range("V1") = "Report" Or Range("a1") = "Report" Then
        X = EssVSetSheetOption(Null, 9, "   -")
    Else
        If Range("a1") = "Bridge" Then
            X = EssVSetSheetOption(Null, 9, "0")
        Else
            X = EssVSetSheetOption(Null, 9, "#missing")
       End If
    End If
    
    If X <> 0 Then
        MsgBox "Error setting #missing label."
    End If
End Sub

Sub ReLogInEssbase()


' Allow to Re login with a different username without close the template.

    application.StatusBar = "Disconnectig all sheets and clearing template variables, please wait..."
    Sheets("Settings").Select
    application.ScreenUpdating = False
    
    'set up a switch on 1 to sopecify re login, not close
    Sheets("Control").Range("SwitchReLogin").Value = 1
        
    'disconnect all sheet
    Call DisconnectAllSheets
    
    'run the clear kitchen stuff macro but check for the switch to do not close the template.
    'Call AAAClearKitchenStuff
    
    'clean the frmLogin texts boxes
    ReqfrmLogin.txtUser = ""
    ReqfrmLogin.txtPwd.Text = ""
    ReqfrmLogin.txtUser.SetFocus
    
    application.StatusBar = False
    
    Sheets("Settings").Select
    
    'call the relogin process again
    Call StartApplication
    
    Sheets("Control").Range("SwitchReLogin").Value = 0
    application.StatusBar = False
End Sub

Sub DisconnectAllSheets()
    Dim sheetCount As Integer
    Dim i As Integer
    
    application.ScreenUpdating = False
    sheetCount = ActiveWorkbook.Sheets.Count
    For i = 2 To sheetCount
        ActiveWorkbook.Sheets(i).Select
        Call DisConnect
    Next i
End Sub
Sub LaunchLink(i As Integer)
Dim j As Integer
Dim MySheet, des As String
Dim Er As Long

On Error GoTo HLinkError
    j = i
    MySheet = ActiveSheet.name
    
    application.ScreenUpdating = False
    
    Sheets("Settings").Select
    
    Select Case j
        Case 1
            Range("HLinkSky").Select
        Case 2
            Range("HLinkOLQR").Select
        Case 3
            Range("HLinkEssbaseFiles").Select
        Case 4
            Range("HLinkVideo").Select
    End Select
        
    Selection.Hyperlinks(1).Follow NewWindow:=False, AddHistory:=False
    'application.WindowState = xlNormal
    
    Sheets(MySheet).Select
    application.CommandBars("Web").Visible = False
    application.ScreenUpdating = True
    
HLinkError:
Er = Err.Number
des = Err.Description
If Er <> 0 Then
    application.StatusBar = False
    If Er = -2146697208 Then
        application.WindowState = xlMaximized
        Sheets(MySheet).Select
        MsgBox "Failure trying to access the Sky Chefs intranet site." & Chr(13) & "Wrong user name or password used. Please try again.", vbExclamation, DbName
    Else
        Sheets(MySheet).Select
        MsgBox "The following error has occurred : " & Er & " Type: " & des & ".", vbCritical, DbName
        End
    End If
End If

End Sub

Sub RetrieveNoMessage()
Dim lngRefreshed As Long
    lngRefreshed = EssVRetrieve(Empty, Empty, 1)
        If lngRefreshed = 0 Then
           'MsgBox "Retrieve successful.", vbInformation, DBNAME
        Else
          If lngRefreshed = 1020010 And ActiveSheet.name = "AirlineList" Then
            MsgBox "This Kitchen has no Airline in the system.", vbExclamation, DbName
          Else
            MsgBox "Retrieve failed.", vbExclamation, DbName
            End
          End If
        End If
End Sub
Sub LockandSend()
Dim lngLocked As Long
Dim lngSendData As Long
Dim lngUnlocked As Long
Dim X As Long


lngLocked = EssVRetrieve(Empty, Empty, 3)
    If lngLocked = 0 Then
       lngSendData = EssVSendData(Empty, Empty)
       If lngSendData = 0 Then
          'MsgBox ("Lock and Send successful, press Ok to continue."), vbInformation, DBNAME
          
        If Sheets("Settings").Range("K3") = "Canada" Then
        
        X = EssVCalculate("Calculate", "CA3YWP", True)
            If X = 0 Then
            'MsgBox ("Calculation complete.")
            Else
            MsgBox ("Calculation failed.")
            End If
        
        ElseIf Sheets("Settings").Range("K3") = "Canada Bank" Then
        X = EssVCalculate("Calculate", "CA3YWP", True)
            If X = 0 Then
            'MsgBox ("Calculation complete.")
            Else
            MsgBox ("Calculation failed.")
            End If
        
        ElseIf Sheets("Settings").Range("K3") = "South Africa" Then
        X = EssVCalculate("Calculate", "SA3YWP", True)
            If X = 0 Then
            'MsgBox ("Calculation complete.")
            Else
            MsgBox ("Calculation failed.")
            End If
       
        ElseIf Sheets("Settings").Range("K3") = "UK With Gazeley" Then
        X = EssVCalculate("Calculate", "UK3YWP", True)
            If X = 0 Then
            'MsgBox ("Calculation complete.")
            Else
            MsgBox ("Calculation failed.")
            End If
       
        ElseIf Sheets("Settings").Range("K3") = "Chile Retail" Then
        X = EssVCalculate("Calculate", "CL3YWP", True)
            If X = 0 Then
            'MsgBox ("Calculation complete.")
            Else
            MsgBox ("Calculation failed.")
            End If
        
        ElseIf Sheets("Settings").Range("K3") = "Chile Bank" Then
         X = EssVCalculate("Calculate", "CL3YWP", True)
            If X = 0 Then
            'MsgBox ("Calculation complete.")
            Else
            MsgBox ("Calculation failed.")
            End If
            
        ElseIf Sheets("Settings").Range("K3") = "Argentina" Then
         X = EssVCalculate("Calculate", "AR3YWP", True)
            If X = 0 Then
            'MsgBox ("Calculation complete.")
            Else
            MsgBox ("Calculation failed.")
            End If
        
            
        ElseIf Sheets("Settings").Range("K3") = "Brazil Southeast" Then
         X = EssVCalculate("Calculate", "BR3YWP", True)
            If X = 0 Then
            'MsgBox ("Calculation complete.")
            Else
            MsgBox ("Calculation failed.")
            End If
            
        ElseIf Sheets("Settings").Range("K3") = "Mexico Retail" Then
         X = EssVCalculate("Calculate", "MX3YWP", True)
            If X = 0 Then
            'MsgBox ("Calculation complete.")
            Else
            MsgBox ("Calculation failed.")
            End If
        
        ElseIf Sheets("Settings").Range("K3") = "Japan (including Discontinued Operations)" Then
         X = EssVCalculate("Calculate", "JP3YWP", True)
            If X = 0 Then
            'MsgBox ("Calculation complete.")
            Else
            MsgBox ("Calculation failed.")
            End If
                
        ElseIf Sheets("Settings").Range("K3") = "China" Then
         X = EssVCalculate("Calculate", "CN3YWP", True)
            If X = 0 Then
            'MsgBox ("Calculation complete.")
            Else
            MsgBox ("Calculation failed.")
            End If
                       
        ElseIf Sheets("Settings").Range("K3") = "INDIA JOINT VENTURE" Then
         X = EssVCalculate("Calculate", "IN3YWP", True)
            If X = 0 Then
            'MsgBox ("Calculation complete.")
            Else
            MsgBox ("Calculation failed.")
            End If
        
         ElseIf Sheets("Settings").Range("K3") = "Walmart Asia Realty" Then
         X = EssVCalculate("Calculate", "WR3YWP", True)
            If X = 0 Then
            'MsgBox ("Calculation complete.")
            Else
            MsgBox ("Calculation failed.")
            End If
        
         ElseIf Sheets("Settings").Range("K3") = "Div35" Then
         X = EssVCalculate("Calculate", "USD3YWP", True)
            If X = 0 Then
            'MsgBox ("Calculation complete.")
            Else
            MsgBox ("Calculation failed.")
            End If
        
        ElseIf Sheets("Settings").Range("K3") = "Div36" Then
         X = EssVCalculate("Calculate", "USD3YWP", True)
            If X = 0 Then
            'MsgBox ("Calculation complete.")
            Else
            MsgBox ("Calculation failed.")
            End If
        
        ElseIf Sheets("Settings").Range("K3") = "Div37" Then
         X = EssVCalculate("Calculate", "USD3YWP", True)
            If X = 0 Then
            'MsgBox ("Calculation complete.")
            Else
            MsgBox ("Calculation failed.")
            End If
        
        
         ElseIf Sheets("Settings").Range("K3") = "EMEACA" Then
         X = EssVCalculate("Calculate", "EO3YWP", True)
            If X = 0 Then
            'MsgBox ("Calculation complete.")
            Else
            MsgBox ("Calculation failed.")
            End If
                 
        ElseIf Sheets("Settings").Range("K3") = "EMEAUK" Then
        X = EssVCalculate("Calculate", "EO3YWP", True)
            If X = 0 Then
            'MsgBox ("Calculation complete.")
            Else
            MsgBox ("Calculation failed.")
            End If
           
        ElseIf Sheets("Settings").Range("K3") = "Latin America Regional Office" Then
        X = EssVCalculate("Calculate", "LO3YWP", True)
            If X = 0 Then
            'MsgBox ("Calculation complete.")
            Else
            MsgBox ("Calculation failed.")
            End If
        
        ElseIf Sheets("Settings").Range("K3") = "Global George" Then
        X = EssVCalculate("Calculate", "GG3YWP", True)
            If X = 0 Then
            'MsgBox ("Calculation complete.")
            Else
            MsgBox ("Calculation failed.")
            End If
        
        ElseIf Sheets("Settings").Range("K3") = "South Bank" Then
        X = EssVCalculate("Calculate", "SB3YWP", True)
            If X = 0 Then
            'MsgBox ("Calculation complete.")
            Else
            MsgBox ("Calculation failed.")
            End If
        
        End If
            
       Else
          MsgBox ("Send failed. Unlocking data.")
          lngUnlocked = EssVUnlock(Empty)
          If lngUnlocked = 0 Then
             MsgBox "Data unlocked. Try again.", vbExclamation, DbName
             End
          Else
             MsgBox "Data not unlocked. Try again.", vbExclamation, DbName
             End
          End If
       End If
    Else
      MsgBox "Lock failed. Cannot send data.", vbCritical, DbName
      application.StatusBar = False
      End
    End If
    application.StatusBar = False
End Sub

Sub Connect()
Dim lngConnected, a, b As Long
Dim strUser As String
Dim strWord As String
Dim strDatabase As String
Dim strAppName As String
Dim strServer

strUser = Range("EssUser")
strWord = Range("EssWord")
strDatabase = Range("EssDB")
strAppName = Range("EssApp")
strServer = Range("EssServer")

        If ActiveSheet.name = ActiveSheet.name = "Control" Then
            Exit Sub
        Else
            a = EssVSetGlobalOption(5, 4)
            lngConnected = EssVConnect(Empty, strUser, strWord, strServer, strAppName, strDatabase)
            If lngConnected = 0 Then
                'MsgBox "You are connected to Essbase.", vbInformation, DBNAME
            End If
        End If
        b = EssVSetGlobalOption(5, 3)
End Sub
Sub DisConnect()
Dim lngDisconnect As Long
Dim Switch As Integer

Switch = Sheets("Control").Range("SwitchReLogin").Value

lngDisconnect = EssVDisconnect(Empty)

If lngDisconnect = 0 Then
  Range("OnLineStatus").Value = "Off-Line"
  If Switch = 0 Then
     'MsgBox ("Disconnect successful."), vbInformation, DBNAME
  Else
    'do nothing
  End If
Else
    Select Case lngDisconnect
        Case -4
          If Switch = 0 Then
            MsgBox "There are no active connections.", vbExclamation, DbName
          Else
            'do nothing
          End If
        Case Is < 0
            MsgBox ("Disconnect failed. Local failure."), vbCritical, DbName
        Case Else
            MsgBox ("Disconnect failed. Server failure."), vbCritical, DbName
    End Select
End If

End Sub

Sub Retrieve()
    Dim lngRefreshed As Long
    
'If the selected sheet is Please Wait, exit sub
If Left(ActiveSheet.name, 4) = "Plea" Then
    Exit Sub
End If
    
    lngRefreshed = EssVRetrieve(Empty, Empty, 1)
        If lngRefreshed = 0 Then
           MsgBox "Retrieve successful.", vbInformation, DbName
        Else
           MsgBox "Retrieve failed.", vbExclamation, DbName
        End If
End Sub

