Attribute VB_Name = "Module3"
Option Explicit

'API function declaration for both 32 and 64bit Excel.
#If VBA7 Then
    Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
        (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#Else
    Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
        (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If
 
Sub DownloadFiles()
                    
    '--------------------------------------------------------------------------------------------------
    'The macro loops through all the URLs (column C) and downloads the files at the specified folder.
    'The given file names (column D) are used to create the full path of the files.
    'If the file is downloaded successfully an OK will appear in column E (otherwise an ERROR value).
    'The code is based on API function URLDownloadToFile, which actually does all the work.
            
    'Written by:    Christos Samaras
    'Date:          02/11/2013
    'Last Update:   28/05/2014
    'e-mail:        xristos.samaras@gmail.com
    'site:          http://www.myengineeringworld.net
    '--------------------------------------------------------------------------------------------------
    
    'Declaring the necessary variables.
    Dim DownloadFolder      As String
    Dim LastRow             As Long
    Dim SpecialChar()       As String
    Dim SpecialCharFound    As Double
    Dim FilePath            As String
    Dim i                   As Long
    Dim j                   As Integer
    Dim Result              As Long
    Dim CountErrors         As Long
    Dim DownloadURL         As String
    Dim LYear As String, LMonth As String, Lday As String
    Dim LYear1 As String, LMonth1 As String, Lday1 As String
    Dim startDate As String, endDate  As String
    Dim MSCIindices As String
    
    'LYear = Format(Now(), "yyyy")
    'j = Val(LYear) - 4
    'LYear1 = CStr(j)
    'LMonth = Format(Now(), "mmm")
    'Lday = Format(Now(), "dd")
    
    LYear = Format(Sheets("Result").Range("C2").Value, "yyyy")
    LMonth = Format(Sheets("Result").Range("C2").Value, "mmm")
    Lday = Format(Sheets("Result").Range("C2").Value, "dd")
    
    LYear1 = Format(Sheets("Result").Range("C4").Value, "yyyy")
    LMonth1 = Format(Sheets("Result").Range("C4").Value, "mmm")
    Lday1 = Format(Sheets("Result").Range("C4").Value, "dd")
    
    startDate = Lday & "%20" & LMonth & ",%20" & LYear
    endDate = Lday1 & "%20" & LMonth1 & ",%20" & LYear1
    MSCIindices = Sheets("MSCI_Index_List").Cells((Sheets("MSCI").Range("K3").Value + 1), 3)
    
    'Disable screen flickering.
    Application.ScreenUpdating = False
    
    'http://www.msci.com/webapp/indexperf/charts?indices=2591,C,36&startDate=23%20Aug,%202010&endDate=22%20Aug,%202014&priceLevel=0&currency=15&frequency=D&scope=R&format=XLS&baseValue=false&site=gimi
    'DownloadURL = "http://www.msci.com/webapp/indexperf/charts?indices=2591,C,36&startDate=23%20Aug,%202010&endDate=22%20Aug,%202014&priceLevel=0&currency=15&frequency=D&scope=R&format=XLS&baseValue=false&site=gimi"
    DownloadURL = "http://www.msci.com/webapp/indexperf/charts?indices=" & MSCIindices & "&startDate=" & startDate & "&endDate=" & endDate & "&priceLevel=0&currency=15&frequency=D&scope=R&format=XLS&baseValue=false&site=gimi"
    
    'An array with special characters that cannot be used for naming a file.
    SpecialChar() = Split("\ / : * ? " & Chr$(34) & " < > |", " ")
    
    'Find the last row and clear the Results column.
    With Sheets("MSCI")
        .Activate
        LastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
        '.Range("E8:E" & LastRow).Value = ""
    End With
    
    'Check if the download folder exists.
    DownloadFolder = Sheets("MSCI").Range("L3").Value
    On Error Resume Next
    If DownloadFolder = "" Or Dir(DownloadFolder, vbDirectory) = vbNullString Then
        MsgBox "The folder's path is incorrect!", vbCritical, "Wrong folder's path"
        Sheets("MSCI").Range("L3").Select
        Exit Sub
    End If
    On Error GoTo 0
               
    'Check if there is at least one URL.
    'If LastRow < 8 Then
    '    MsgBox "You did't enter a URL!", vbCritical, "No URL"
    '    Sheets("MSCI").Range("C8").Select
    '    Exit Sub
    'End If
    
    'Add the backslash if doesn't exist.
    If Right(DownloadFolder, 1) <> "\" Then
        DownloadFolder = DownloadFolder & "\"
    End If
            
    'Counting the number of files that will not be downloaded.
    CountErrors = 0
    
    'Save the internet files at the specified folder of your hard disk.
    On Error Resume Next
    'For i = 8 To LastRow
    i = 3
    
        'Use the given file name.
        If Not Cells(i, 13).Value = vbNullString Then
            
            'Get the given file name.
            FilePath = Cells(i, 13).Value
            
            'Check if the file path contains a special/illegal character.
            For j = LBound(SpecialChar) To UBound(SpecialChar)
                SpecialCharFound = WorksheetFunction.Find(SpecialChar(j), FilePath)
                'If an illegal character is found substitute it with a "-" character.
                If SpecialCharFound > 0 Then
                    FilePath = WorksheetFunction.Substitute(FilePath, SpecialChar(j), "-")
                End If
            Next j
            
            'Create the final file path.
            FilePath = DownloadFolder & FilePath
            
            'Check if the file path exceeds the maximum allowable characters.
            If Len(FilePath) > 255 Then
                Cells(i, 14).Value = "ERROR"
            End If
        Else
            'Empty file name.
            Cells(i, 14).Value = "ERROR"
        End If
        
        'If the file path is valid save the file to the selected folder.
        If UCase(Cells(i, 14).Value) <> "ERROR" Then
        
            'Save the files to the selected folder.
            'Result = URLDownloadToFile(0, Cells(i, 3).Value, FilePath, 0, 0)
            Result = URLDownloadToFile(0, DownloadURL, FilePath, 0, 0)
            
            'Check if the file downloaded successfully  and if it exists.
            If Result = 0 And Not Dir(FilePath, vbDirectory) = vbNullString Then
                'Success!
                Cells(i, 14).Value = "OK"
            Else
                'Error!
                Cells(i, 14).Value = "ERROR"
                CountErrors = CountErrors + 1
            End If
            
        End If
        
    'Next i
    On Error GoTo 0
    
    'Inform the user that macro finished successfully or with errors and enable the screen.
    'If CountErrors = 0 Then
    '    'Success!
    '    If LastRow - 7 = 1 Then
    '        Application.ScreenUpdating = True
    '        MsgBox "The file was successfully downloaded!", vbInformation, "Done"
    '    Else
    '        Application.ScreenUpdating = True
    '        MsgBox LastRow - 7 & " files were successfully downloaded!", vbInformation, "Done"
    '    End If
    'Else
    '    'Error!
    '    If CountErrors = 1 Then
    '        Application.ScreenUpdating = True
    '        MsgBox "There was an error in one of the files!", vbCritical, "Error"
    '    Else
    '        Application.ScreenUpdating = True
    '        MsgBox "There was an error in " & CountErrors & " files!", vbCritical, "Error"
    '    End If
    'End If
    
End Sub


