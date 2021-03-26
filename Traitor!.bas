Attribute VB_Name = "Module1"
' Traitor! - An experimental password cracker for Excel files, written in VBA
'
' Copyright (c) 2020-2021 Craig Jackson
'
' DISCLAIMER - Traitor! was created for authorized personal and professional testing. Using it to attack targets without prior mutual consent is illegal.
' It is the end user’s responsibility to obey all applicable local, state, and federal laws. The author(s) assume no liability and are not responsible
' for any misuse or damage caused by this tool.

Public formStart As Boolean, endAll As Boolean
Public tryCount As Integer
Public Rng As Range
Public pwList As Worksheet, aWks As Worksheet

Sub main()

' Attempts to crack the encryption password on a protected Microsoft Excel
' workbook using a plaintext wordlist (one password per line), or by creating
' a sequential list of brute force password candidates.

' NOTE: The Microsoft Scripting Runtime reference is required before this
' script can be run. You can add that in the VBA editor under Tools > References...

Application.ScreenUpdating = False

Dim chkCaps As Boolean, chkLows As Boolean, chkNums As Boolean, chkSpecs As Boolean
Dim impWrds As Boolean, delWrds As Boolean, pwdResult As Boolean, wksCheck As Boolean
Dim startTm As Double
Dim fso As FileSystemObject
Dim minLen As Integer, maxLen As Integer, lastClmn As Integer, maxTries As Integer
Dim thisPwd As String, charPool As String
Dim attType As String, tgtType As String, tgtPath As String, pwdPath As String, pwdTest As String
Dim runTm As String
Dim pwdLines As TextStream

tryCount = 0
thisPwd = ""
startTm = Timer
endAll = False
pwdFound = False
wksCheck = False
Set fso = New FileSystemObject

If formStart = False Then

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''' CHANGE USER OPTIONS AS DESIRED ''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Attack Type options:
    ' - Dict = Dictionary attack using an external wordlist. This should be a txt file with one password per line
    ' - Brute = Traitor! will procedurally generate passwords and use them against an excel workbook
    ' - minLen = The minimum password length that should be used against the target
    ' - maxLen = The maximum password length that should be used against the target
    ' - maxTries = The maximum number of guesses Traitor! will make regardless of password length settings

' Target Type options:
    ' - File = Attempt to open a password-protected (encrypted) XLS or XLSX file
    ' - Workbook = Attempt to unlock a workbook with the "Protect Workbook" password set
    ' - Worksheet = Attempt to unlock a worksheet with the "Protect Worksheet" password set
    
' Target Path options:
    ' - File target path (tgtPath) requires the fully-qualified path to the target file, including file extension
    ' - Workbook target path can use basic workbook name (e.g., "Book1.xlsx"), but the workbook needs to be open
    ' - Worksheet target path requires the worksheet name
    ' - For workbook and worksheet target types, leave the tgtPath variable blank (i.e., "") to attack the active workbook or worksheet
    
' Dictionary attack options:
    ' - impWrds = "True" attempts to import wordlist to cells within the workbook, "False" grabs words from TextStream object in memory
    ' - delWrds = "True" automatically deletes the pwLines spreadsheet when complete if words were imported, "False" leaves it
    
' Brute Force attack options:
    ' - chkCaps = "True" includes capital letters in the character pool for brute forcing, "False" does not
    ' - chkLows = "True" includes lowercase letters in the character pool for brute forcing, "False" does not
    ' - chkNums = "True" includes numbers in the character pool for brute forcing, "False" does not
    ' - chkSpecs = "True" includes special characters / symbols in the character pool for brute forcing, "False" does not

    tgtType = "File" ' Target Type
    attType = "Brute" ' Attack Type
    tgtPath = "C:\Users\cjackson\Desktop\ThisisaTest.xlsx" ' Target Path
    pwdPath = "C:\Users\cjackson\Desktop\testpwds2.txt" ' Path to password dictionary file; fully-qualify path
    impWrds = True ' Import wordlist or run from TextStream object?
    delWrds = False ' Automatically delete pwList sheet when complete?
    chkCaps = False ' Include caps if brute forcing
    chkLows = True ' Include lowers if brute forcing
    chkNums = False ' Include numbers if brute forcing
    chkSpecs = False ' Include special characters if brute forcing
    minLen = 1 ' Minimum password length (min allowed: 1, default: 1)
    maxLen = 15 ' Maximum password length (max allowed: 50, default: 15)
    maxTries = 0 ' Maximum number of password guesses before exit (0 for infinite, default: 0)
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Else

' Otherwise, pull configuration from user settings on ControlForm

    With ControlForm
        If .FileTarget.Value = True Then tgtType = "File" ' Target type
        If .WorkbookTarget.Value = True Then tgtType = "Workbook"
        If .WorksheetTarget.Value = True Then tgtType = "Worksheet"
        If .DictAtt.Value = True Then attType = "Dict" ' Attack Type
        If .BrutAtt.Value = True Then attType = "Brute"
        
        tgtPath = .TargetName.Value ' Target Path
        pwdPath = .DictName.Value ' Path to password dictionary file; fully-quality path
        
        If .minLen.Value = "" Then ' Minimum password length (default: 1)
            minLen = 1
        ElseIf .minLen >= 1 Or .minLen < 50 Then
            minLen = .minLen.Value
        End If
        
        If .maxLen.Value = "" Then ' Maximum password length (default: 15)
            maxLen = 15
        ElseIf .maxLen > 2 Or .maxLen <= 50 Then
            maxLen = .maxLen.Value
        End If
        
        If .maxTries.Value = "" Then ' Maximum number of password guesses (default: 0 for infinite)
            maxTries = 0
        Else
            maxTries = .maxTries.Value
        End If
            
        If .DictOpt1.Value = True Then ' Import wordlist or run from TextStream object?
            impWrds = True
        Else
            impWrds = False
        End If
        
        If .DictOpt2.Value = True Then ' Automatically delete pwList sheet when complete?
            delWrds = True
        Else
            delWrds = False
        End If
        
        If .chkCaps = True Then ' Include caps if brute forcing
            chkCaps = True
        Else
            chkCaps = False
        End If
        
        If .chkLows = True Then ' Include lowers if brute forcing
            chkLows = True
        Else
            chkLows = False
        End If
        
        If .chkNums = True Then ' Include lowers if brute forcing
            chkNums = True
        Else
            chkNums = False
        End If
        
        If .chkSpecs = True Then ' Include lowers if brute forcing
            chkSpecs = True
        Else
            chkSpecs = False
        End If
    End With
    
End If

' Check to make sure the target specified by the user exists

If tgtType = "Workbook" And tgtPath = "" Then
    tgtPath = ActiveWorkbook.Name ' Default target to active workbook
ElseIf tgtType = "Worksheet" And tgtPath = "" Then
    tgtPath = activeworksheet.Name ' Default target to active worksheet
ElseIf checkTgt(tgtType, tgtPath) = False Then
    MsgBox "[ERROR] Doesn't look like target " & LCase(tgtType) & " called " & tgtPath & " exists."
    Exit Sub
End If

' Check to make sure the user has passed allowed values for minLen and maxLen variables

If Not IsNumeric(maxLen) Or Not IsNumeric(minLen) Then ' User passed something other than numbers for max and min length

    MsgBox "Please pass only whole numbers for the minLen and maxLen variables to proceed.", vbExclamation, "Invalid selection"
    Exit Sub

End If

If maxLen < minLen Then ' User passed a max length value that is less than the min legnth value
    
    MsgBox "Please be sure that the maxLen variable value is greater than or equal to the minLen variable value.", vbExclamation, "Invalid selection"
    Exit Sub
    
End If
    
If Not minLen >= 1 And minLen <= 50 Then ' User passed a min length value that does not fall in the allowed range
    
    MsgBox "Please set the minLen variable to a whole number between 1 and 50 to proceed.", vbExclamation, "Invalid selection"
    Exit Sub
    
End If

If Not maxLen >= 2 And minLen <= 50 Then ' User passed a max length value that does not fall in the allowed range
    
    MsgBox "Please set the maxLen variable to a whole number between 2 and 50 to proceed.", vbExclamation, "Invalid selection"
    Exit Sub
    
End If

' Determine whether dictionary or brute force attack is taking place

If attType = "Brute" Then

    ' Confirm that a character selection has been made
    
    If chkCaps = False And chkLows = False And chkNums = False And chkSpecs = False Then
        
        MsgBox "Please be sure to pass a 'True' value for at least one of the brute force character types " & _
            "in the 'Change User Options as Desired' section above.", vbExclamation, "Missing selection"
        Exit Sub
    
    End If
    
    ' Create the character pool for the attack
    
    charPool = genCharPool(chkCaps, chkLows, chkNums, chkSpecs)
    
    ' Execute the bruting process
    
    Do
    
        pwdTest = genPwd(pwdTest, minLen, maxLen, charPool)
        
        If endAll = False Then
            pwdResult = tryPassword(pwdTest, tgtType, tgtPath) ' Try the current password
            If tryCount = maxTries Then Exit Do ' User-defined password attempt limit reached
        Else
            Exit Do ' Break; all potential passwords exhausted
        End If
        
    Loop Until pwdResult = True

ElseIf attType = "Dict" Then

    ' Check to make sure the wordlist file specified by the user exists
    
    If Dir(pwdPath) = "" Then
        MsgBox "[ERROR] Doesn't look like your password list exists at " & pwdPath & "."
        Exit Sub
    End If
    
    ' Open password dictionary file as a TextStream object
    
    Set pwdLines = fso.OpenTextFile(pwdPath, ForReading, False)
    
    ' Try to open the encrypted workbook by cycling through the imported passwords
    
    If impWrds = True Then
    
        ' First create (if necessary) and select new worksheet
        
        For Each aWks In ActiveWorkbook.Worksheets
            If aWks.Name = "PwList" Then wksCheck = True ' pwList already exists
        Next aWks
    
        If wksCheck = False Then
            On Error Resume Next
            Sheets.Add(Before:=Sheets(1)).Name = "PwList" ' Create pwList if needed
        End If
        
        ' Check to make sure the worksheet add was successful
        
        If Not Err Is Nothing Then
            MsgBox "Creation of the PwList spreadsheet failed. Please remove the 'Import Dictionary Words' selection and try again.", vbExclamation, "Process error"
            Exit Sub
        End If
        
        Set pwList = Worksheets("PwList")
        pwList.Select
    
        ' Then import the words into the pwList spreadsheet,
        ' then close the TextStream object
    
        Call importWords(pwdLines, minLen, maxLen)
        pwdLines.Close
    
        ' Get the rightmost column that isn't empty. If the user has managed to fill
        ' all columns out to the last column Excel allows (XFD or 16384) with passwords,
        ' that column will be the rightmost. Of course, Excel will likely have crashed
        ' at that point anyway, so...
    
        If Not IsEmpty(pwList.Range("XFD1")) Then
            lastClmn = pwList.Range("XFD1").Column
        Else
            lastClmn = pwList.Range("XFD1").End(xlToLeft).Column
        End If
        
        ' Cycle through the imported passwords
        
        For Each Rng In Range(pwList.Range("A1"), pwList.Cells(1048576, lastClmn))
            If Not IsEmpty(Rng) Then
                pwdTest = Rng.Value
                pwdResult = tryPassword(pwdTest, tgtType, tgtPath)
                If tryCount = maxTries Then Exit For ' User-defined password attempt limit reached
            End If
            
            If pwdResult = True Then
                Exit For ' Break loop on success
            End If
        Next Rng
        
        ' Remove the password list spreadsheet that was created
        ' if the user wants to
        
        If delWrds = True Then
            Application.DisplayAlerts = False ' Turn off confirmations
            pwList.Delete ' Delete the password list worksheet
            Application.DisplayAlerts = True
        End If
    
    Else
    
        ' Run password guesses directly from TextStream file in memory
        
        Do Until pwdLines.AtEndOfStream
        
            pwdTest = pwdLines.ReadLine
            
            ' Confirm that the currently tested password meets the necessary
            ' length requirements and test it if it does.
            
            If Len(pwdTest) >= minLen And Len(pwdTest) <= maxLen Then
                
                pwdResult = tryPassword(pwdTest, tgtType, tgtPath)
            
                ' If the password is found or the user-defined max number
                ' of tries is reached, exit.
            
                If pwdResult = True Or tryCount = maxTries Then
                    pwdLines.Close
                    Exit Do ' Break loop on success
                End If
                
            End If
        
        Loop
    
    End If

End If

' Close out, notify user

runTm = Format((Timer - startTm) / 86400, "hh:mm:ss")

If pwdResult = False Then
    MsgBox "Failed to get password after " & tryCount & " tries in " & runTm & "."
Else
    MsgBox "Success after " & tryCount & " tries in " & runTm & "!" & vbNewLine & "Password is: " & pwdTest
End If

Application.ScreenUpdating = True

End Sub

Sub importWords(pwdLines As TextStream, minLen As Integer, maxLen As Integer)

' Imports a list of words from a text file and lists them one-per-row
' in a pre-created worksheet called "PwList"

Dim fileEnd As Boolean
Dim startTm As Double, pwdCount As Double
Dim colNum As Integer
Dim rowNum As Long
Dim thisPwd As String

Set pwList = Worksheets("pwList")

fileEnd = False
pwdCount = 0
rowNum = 1
colNum = 1
startTm = Timer

Do

    rowNum = 1 ' Set top row on each new column

    Do
    
        Set Rng = pwList.Cells(rowNum, colNum) ' Set current import target cell
        thisPwd = pwdLines.ReadLine ' Set current password attempt
        
        ' Check to make sure that the length of the current password is
        ' Greater than or equal to the minimum length, and shorter than or
        ' equal to the maximum length
        
        If Len(thisPwd) >= minLen And Len(thisPwd) <= maxLen Then
        
            With Rng
                .NumberFormat = "@" ' Set cell format to "Text"
                .Value = thisPwd
            End With
        
            pwdCount = pwdCount + 1 ' Increment imported password count
            rowNum = rowNum + 1 ' Advance row
            
        End If
        
        ' This code block prevents an infinite loop
        ' by breaking out of both Do loops when the
        ' last line of the text stream is reached.
        
        If pwdLines.AtEndOfStream = True Then
            fileEnd = True
            Exit Do
        End If
    
    Loop Until Rng.Row = 1048576 ' Break when max row is hit

    If fileEnd = True Then Exit Do ' Import completed

    colNum = colNum + 1 ' Advance column
    
Loop

End Sub

Function genCharPool(chkCaps As Boolean, chkLows As Boolean, chkNums As Boolean, chkSpecs As Boolean) As String

' Creates the character pool for use with brute force attachs

Dim charCaps As String, charLows As String, charNums As String, charSyms As String, charPool As String

charLtrs = "abcdefghijklmnopqrstuvwxyz"
charCaps = UCase(charLtrs)
charLows = LCase(charLtrs)
charNums = "1234567890"
charSyms = "!@#$%^&*()_+[]\<>?,./"
charPool = ""

' Creates the character pool based on the users selections

If chkCaps = True Then genCharPool = genCharPool + charCaps
If chkLows = True Then genCharPool = genCharPool + charLows
If chkNums = True Then genCharPool = genCharPool + charNums
If chkSpecs = True Then genCharPool = genCharPool + charSyms

End Function

Function genPwd(thisPwd As String, minLen As Integer, maxLen As Integer, chrPool As String) As String

' Sequentially creates passwords to use with a brute force attack

Dim newPwd As Boolean, cyclePwd As Boolean
Dim iterIdx As Integer, iterPool As Integer, currLen As Integer
Dim oldChr As String, newChr As String

cyclePwd = False
newPwd = False
pwdCount = 0

genPwd = thisPwd ' Set function variable value to current password

If genPwd = "" Then

    ' Create the initial password based on the minLen and charPool variables
    
    Do
        genPwd = genPwd & Left(chrPool, 1)
    Loop Until Len(genPwd) = minLen
    
Else

    ' Iterate the existing password
    
    iterIdx = Len(genPwd) ' Start iteration at end of current password and work back
    currLen = Len(genPwd) ' Get length of the current password
    
    Do
    
        oldChr = Mid(genPwd, iterIdx, 1) ' Get the character at the current iterator position
        iterPool = InStr(chrPool, oldChr) ' Find the index of the character being replaced in the character pool
        
        If iterPool = Len(chrPool) Then
        
            ' The character at this iterator index has reached the end of the
            ' available characters in the character pool. The iterator index will
            ' need to be decreased by one to move to the previous character in the
            ' password, which will now be iterated.
            
            If iterIdx = 1 Then
            
                ' When the character iterator has cycled back to the first position in
                ' the password and the character in that position is the last character
                ' in the charPool variable, one of two things happens...
                
                If Len(genPwd) = maxLen Then
                
                    ' If an additional character can't be added to the existing password
                    ' because the max length has been reached, all password options have been
                    ' exhausted. Set the endAll variable to True and exit the function.
                    
                    endAll = True
                    Exit Function
                
                Else
                
                    ' If there's still room to grow the password, add an additional
                    ' character to the end and cycle all characters back to the first
                    ' character of chrPool. Once completed, exit function
                    
                    genPwd = ""
                    
                    Do
                        genPwd = genPwd & Left(chrPool, 1)
                    Loop Until Len(genPwd) = currLen + 1
                    
                    Exit Function
                    
                End If
                
            Else
            
                ' cyclePwd tells the sub that all characters after this index in the
                ' string will need to be cycled back to the first character of the chrPool.
            
                iterIdx = iterIdx - 1
                cyclePwd = True
                
            End If
            
        Else
        
            newChr = Mid(chrPool, iterPool + 1, 1) ' Get the next character in the character pool
            newPwd = True ' Break; the new password is ready to be built
            
        End If
        
    Loop Until newPwd = True
    
    If cyclePwd = True Then

        ' A character midway through the current password is being iterated. All characters
        ' beyond it can be removed and replaced with the first character of chrPool.
        
        genPwd = Left(genPwd, iterIdx - 1) & newChr
        
        Do Until Len(genPwd) = currLen Or Len(genPwd) = maxLen
            genPwd = genPwd & Left(chrPool, 1)
        Loop
    
    Else

        ' Only the end character is being iterated. No other characters need to be replaced.
    
        genPwd = Left(genPwd, Len(genPwd) - 1) & newChr ' Replace old character with new
    
    End If

End If

End Function

Function tryPassword(pwdTest As String, tgtType As String, tgtPath As String) As Boolean

' Tests the current password on the target workbook or worksheet

Dim tgtWb As Workbook

tryPassword = False

' Try to open or unlock target using the current password

On Error Resume Next

    tryCount = tryCount + 1

    Select Case tgtType
    
        Case Is = "File"

            Set tgtWb = Workbooks.Open(tgtPath, Password:=pwdTest)
            
            If InStr(1, Err.Description, "The password you supplied is not correct") = False Then
                
                ' If you get here you guessed the right password!
                
                Workbooks(tgtWb.Name).Close savechanges:=False
                tryPassword = True
                
            End If
            
        Case Is = "Workbook"
        
            ActiveWorkbook.Unprotect Password:=pwdTest
     
            If InStr(1, Err.Description, "The password you supplied is not correct") = False Then
                
                ' If you get here you guessed the right password!
                
                tryPassword = True
                
            End If
        
        Case Is = "Worksheet"
        
            ActiveSheet.Unprotect Password:=pwdTest
     
            If InStr(1, Err.Description, "The password you supplied is not correct") = False Then
                
                ' If you get here you guessed the right password!
                
                tryPassword = True
                
            End If
        
    End Select

On Error GoTo -1

End Function

Function checkTgt(tgtType As String, tgtPath As String) As Boolean

' Checks to make sure the target passed by the user is valid

Dim wbTest As Workbook

checkTgt = False

Select Case tgtType

    Case Is = "File"
    
        If Not Dir(tgtPath) = "" Then
            checkTgt = True ' Valid target path
        End If
    
    Case Is = "Workbook"
    
        On Error Resume Next
            Set wbTest = Application.Workbooks.Item(tgtPath)
            If Not Err.Description = "Subscript out of range" Then checkTgt = True ' Valid target path
        On Error GoTo -1
    
    Case Is = "Worksheet"
    
        For Each aWks In ActiveWorkbook.Worksheets
            If aWks.Name = tgtPath Then checkTgt = True ' Valid target path
        Next aWks

End Select

End Function

Sub openControlForm()

' Opens the ControlForm by clicking the start button on the info page

With ControlForm
    .StartUpPosition = 0
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    .Show (0)
End With

End Sub
