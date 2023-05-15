Dim objShell, waitTime, websiteURL, userChoice
Set objShell = CreateObject("WScript.Shell")
waitTime = 5200 ' Delay in milliseconds (5.2 seconds)

' Display the welcome message
MsgBox "Welcome to CasonPlayz Web Tester Script/App!" & vbCrLf & "Which one would you like to choose?" & vbCrLf & "1. Just ping (aka quick test)" & vbCrLf & "2. Ping and curl (default test)" & vbCrLf & "3. Port viewer and manager (coming soon)", vbInformation, "Web Tester"

' Prompt the user to choose an option
userChoice = InputBox("Choose an option:", "Options")

' Process user's choice
Select Case userChoice
    Case "1"
        ' Just ping (quick test)
        websiteURL = InputBox("Enter the website URL to ping:", "Website URL Input")
        Dim pingCommand
        pingCommand = "ping -n 1 " & websiteURL
        Dim pingResult
        pingResult = objShell.Run(pingCommand, 0, True)
        
        ' Check the ping result
        If pingResult = 0 Then
            ' Ping successful
            MsgBox "Ping successful for " & websiteURL, vbInformation, "Ping Result"
        Else
            ' Ping failed
            MsgBox "Ping failed for " & websiteURL, vbCritical, "Ping Result"
        End If

    Case "2", "" ' Default test
        ' Ping and curl
        websiteURL = InputBox("Enter the website URL to ping and curl:", "Website URL Input")
        Dim pingCommand2 ' Rename the variable
        pingCommand2 = "ping -n 1 " & websiteURL
        Dim pingResult2
        pingResult2 = objShell.Run(pingCommand2, 0, True)
        
        ' Check the ping result
        If pingResult2 = 0 Then
            ' Ping successful
            MsgBox "Ping successful for " & websiteURL & ". Waiting for " & waitTime & " milliseconds before executing the curl command.", vbInformation, "Ping Result"
            WScript.Sleep waitTime
            
            ' Execute the curl command
            Dim curlCommand
            curlCommand = "curl " & websiteURL
            objShell.Run curlCommand
            
            ' Display a success message
            MsgBox "Curl command executed successfully for " & websiteURL, vbInformation, "Curl Result"
        Else
            ' Ping failed
            MsgBox "Ping failed for " & websiteURL, vbCritical, "Ping Result"
        End If

    Case "3"
        ' Port viewer and manager (coming soon)
        MsgBox "Port viewer and manager feature is coming soon.", vbInformation, "Feature Coming Soon"

End Select
