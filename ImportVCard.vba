Option Explicit

'------------------------------------------------------------------------------
' Name:        ImportVCard
' Author:      Ralf
' Purpose:     Opens all vCard files from a local folder in Outlook, saves the
'              imported item, and closes the Inspector window afterwards.
'
' Description:
' This macro loops through all files in C:\VCARDS, opens each file via Outlook,
' waits until an Inspector window is available, saves the currently opened item,
' and then closes it.
'
' Requirements:
' - Microsoft Outlook installed
' - Folder C:\VCARDS must exist
' - vCard files must be stored in C:\VCARDS
'
' Notes:
' - The macro uses the Outlook Inspectors collection to detect opened items.
' - The imported item is accessed through CurrentItem.
' - The Inspector is closed using olDiscard after the item has been saved.
'------------------------------------------------------------------------------

Sub OpenSaveVCard()

    Dim fso As Object
    Dim fsDir As Object
    Dim fsFile As Object
    
    Dim objOL As Object
    Dim colInsp As Object
    Dim objWSHShell As Object
    
    Dim strVCName As String

    On Error Resume Next

    ' Create FileSystemObject and access the vCard folder
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fsDir = fso.GetFolder("C:\VCARDS")

    ' Loop through all files in the folder
    For Each fsFile In fsDir.Files

        ' Build the full path to the current vCard file
        strVCName = "C:\VCARDS\" & fsFile.Name

        ' Create Outlook application object
        Set objOL = CreateObject("Outlook.Application")
        Set colInsp = objOL.Inspectors

        ' Only continue if no Inspector window is currently open
        If colInsp.Count = 0 Then

            ' Open the vCard file using the Windows shell
            Set objWSHShell = CreateObject("WScript.Shell")
            objWSHShell.Run Chr(34) & strVCName & Chr(34)

            DoEvents
            Set colInsp = objOL.Inspectors

            If Err = 0 Then

                ' Wait until the Inspector window is available
                Do Until colInsp.Count = 1
                    DoEvents
                Loop

                ' Save the imported Outlook item
                colInsp.Item(1).CurrentItem.Save

                ' Close the Inspector and discard further changes
                colInsp.Item(1).Close olDiscard

                ' Release object references
                Set colInsp = Nothing
                Set objOL = Nothing
                Set objWSHShell = Nothing

            End If
        End If

    Next fsFile

    ' Cleanup
    Set fsFile = Nothing
    Set fsDir = Nothing
    Set fso = Nothing

End Sub