Attribute VB_Name = "Module2"
Option Explicit
    
    '---------------------------------------------------
    'This module contains some auxiliary subs.
   
    'Written by:    Christos Samaras
    'Date:          02/11/2013
    'e-mail:        xristos.samaras@gmail.com
    'site:          http://www.myengineeringworld.net
    '---------------------------------------------------
    
Sub FolderSelection()
    
    'Shows the folder picker dialog in order the user to select the folder
    'in which the downloaded files will be saved.
    
    Dim FoldersPath As String
    
    'Show the folder picker dialog.
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select a folder to save your files..."
        .Show
            If .SelectedItems.Count = 0 Then
                Sheets("MSCI").Range("L3").Value = "-"
                MsgBox "You did't select a folder!", vbExclamation, "Canceled"
                Exit Sub
            Else
                FoldersPath = .SelectedItems(1)
            End If
    End With
    
    'Pass the folder's path to the cell.
    Sheets("MSCI").Range("L3").Value = FoldersPath
    
End Sub

Sub Clear()
    
    'Clears the URLs, the result column and the folder's path.
            
    'Dim LastRow As Long
       
    'Find the last row.
    'With Sheets("MSCI")
    '    .Activate
    '    LastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
    'End With
    
    'Clear the ranges.
    'If LastRow > 8 Then
    '    Sheets("MSCI").Range(Cells(8, 3), Cells(LastRow, 5)).Value = ""
    'End If
    'Sheets("MSCI").Range("L3").Value = ""
    'Sheets("MSCI").Range("L3").Select
    
End Sub

