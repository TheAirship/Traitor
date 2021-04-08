VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ControlForm
   Caption         =   "Traitor! Control Form"
   ClientHeight    =   8640.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13335
   OleObjectBlob   =   "ControlForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ControlForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Traitor! - An experimental Excel-based password cracker for Excel files.
'
' DISCLAIMER - Traitor! was created for authorized personal and professional testing.
' Using it to attack targets without prior mutual consent is illegal.
' It is the end user�s responsibility to obey all applicable local, state, and
' federal laws. The author(s) assume no liability and are not responsible
' for any misuse or damage caused by this tool.
'
' Copyright (c) 2020-2021 Craig Jackson
'
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
' http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.

Private Sub openControlForm()

' START HERE: Use this sub to open the ControlForm manually by selecting
' anywhere in this procedure and pressing the F5 key. You may also select
' Run > Run Sub/UserForm from the menu bar at the top.

With ControlForm
    .StartUpPosition = 0
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    .Show
End With

End Sub

Private Sub BrutAtt_Click()

' Changes control properties on the ControlForm
' based on selection of the Brute Force attack type

If BrutAtt.Value = True Then
    With DictName
        .Value = "N/A"
        .Enabled = False
    End With

    DictBrowse.Enabled = False
    chkCaps.Enabled = True
    chkLows.Enabled = True
    chkNums.Enabled = True
    chkSpecs.Enabled = True

    With DictOpt1
        .Enabled = False
        .Value = False
    End With
    With DictOpt2
        .Enabled = False
        .Value = False
    End With
End If

End Sub

Private Sub GoButton_Click()

' Check to make sure the necessary selections have been made
' starting with ensuring that a target type was selected

If FileTarget.Value = False And WorkbookTarget.Value = False And WorksheetTarget.Value = False Then

    MsgBox "Please select a target type (Step 1) to initiate a password attack.", vbExclamation, "Missing selection"
    Exit Sub

End If

' Check to make sure an attack type was selected

If DictAtt.Value = False And BrutAtt.Value = False Then

    MsgBox "Please select an attack type (Step 2) to initiate a password attack.", vbExclamation, "Missing selection"
    Exit Sub

End If

' Check to make sure a file path was entered if the file target type was selected

If FileTarget.Value = True And TargetName.Value = "" Then

    MsgBox "Please enter the full path to a target Excel file (Step 3) when using the File target type. You can use the browse button to locate a target Excel file.", vbExclamation, "Missing target"
    Exit Sub

End If

' Check to make sure the path to a dictionary file has been entered when using the Dictionary attack type

If DictAtt.Value = True And DictName.Value = "" Then

    MsgBox "Please enter the full path to a wordlist file (Step 4). You can use the browse button to locate the file.", vbExclamation, "Missing wordlist"
    Exit Sub

End If

' Check to make sure at least one character option has been selected with the Brute Force attack type

If BrutAtt.Value = True And (chkCaps.Value = False And chkLows.Value = False And chkNums.Value = False And chkSpecs.Value = False) Then

    MsgBox "Please select at least one character type for Traitor! to use with a brute force attack.", vbExclamation, "Missing selection"
    Exit Sub

End If

formStart = True
Call main

End Sub

Private Sub CancelButton_Click()

Unload Me

End Sub

Private Sub DictAtt_Click()

' Changes control properties on the ControlForm
' based on selection of the Dictionary attack type

If DictAtt.Value = True Then
    With DictName
        .Enabled = True
        .Value = ""
    End With

    DictBrowse.Enabled = True
    DictOpt1.Enabled = True
    DictOpt2.Enabled = True

    With chkCaps
        .Enabled = False
        .Value = False
    End With
    With chkLows
        .Enabled = False
        .Value = False
    End With
    With chkNums
        .Enabled = False
        .Value = False
    End With
    With chkSpecs
        .Enabled = False
        .Value = False
    End With
End If

End Sub

Private Sub DictBrowse_Click()

' Opens the file browse window for the wordlist file

Dim DictPath As String

DictPath = Application.GetOpenFilename

If Not DictPath = "False" Then
    DictName.Value = DictPath
End If

End Sub

Private Sub FileTarget_Click()

' Selects the file target type and enables
' the necessary controls

TargetName.Enabled = True
TargetBrowse.Enabled = True

End Sub

Private Sub TargetBrowse_Click()

' Opens the file browse window for the target file

Dim tgtPath As String

tgtPath = Application.GetOpenFilename

If Not tgtPath = "False" Then
    TargetName.Value = tgtPath
End If

End Sub

Private Sub WorkbookTarget_Click()

' Selects the workbook target type and enables
' the necessary controls

TargetBrowse.Enabled = False

End Sub

Private Sub WorksheetTarget_Click()

' Selects the worksheet target type and enables
' the necessary controls

TargetBrowse.Enabled = False

End Sub
