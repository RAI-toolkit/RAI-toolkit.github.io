Sub ExportRowsAsTextFiles()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowNum As Long
    Dim rowRange As Range
    Dim cellValue As String
    Dim filePath As String
    Dim cell As Range
    
    Set ws = ThisWorkbook.ActiveSheet ' Change to the appropriate sheet if needed
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' Assumes data starts in column A
    
    ' Loop through each row
    For rowNum = 2 To lastRow
        ' Set the range for the current row
        'Set rowRange = ws.Range("A" & rowNum & ":" & ws.Cells(rowNum, ws.Columns.Count).End(xlToLeft).Address)
        
        ' Initialize cellValue as an empty string
        'cellValue = ""
        
        ' Loop through each cell in the row
        'For Each cell In rowRange
            ' Check the data type of the cell value
            'Select Case True
                'Case IsNumeric(cell.Value) ' Numeric value
                    'cellValue = cellValue & CStr(cell.Value) & ","
                'Case IsDate(cell.Value) ' Date value
                    'cellValue = cellValue & Format(cell.Value, "dd-mm-yyyy") & ","
                'Case Else ' Text value or other types
                    'cellValue = cellValue & CStr(cell.Value) & ","
            'End Select
       ' Next cell
        
        ' Remove the trailing comma
        'cellValue = Left(cellValue, Len(cellValue) - 1)
        
        ' Define the file path for the text file (change as needed)
         'Currentlyfiles stored in Library/Containers/Microsoft Excel/ Data
        
        cellCat = Range("V" & rowNum).Value
        cellType = Range("B" & rowNum).Value
                
        RootFolder = ":_cat_" & cellCat & ":"
        'RootFolder = ""
        'RootFolder = Replace(ThisWorkbook.Path, "/", ":")
        
        cellTitle = Range("A" & rowNum).Value
        'Remove special characers from filename
        FileName = Replace(cellTitle, " ", "-")
        FileName = Replace(FileName, ":", "-")
        FileName = Replace(FileName, "?", "-")
        FileName = Left(FileName, 40)
            
        'Append post date to filename
        PostDate = Range("R" & rowNum).Value
        PostDate = Format(PostDate, "yyyy-mm-dd")
            
        postPath = ":_posts:" & PostDate & "-" & FileName & ".md"
                
        'Export the row as a text file
        Open postPath For Output As #1
            Print #1, "---"
            Print #1, "title:  """; cellTitle; """  "
            cellDescription = Range("C" & rowNum).Value
            Print #1, "excerpt:  """; Left(cellDescription, 240); " (...)""  "
                  
            Print #1, "header:"
            'Print #1, "  image: /assets/images/RAI_toolkit/"; Left(cellType, 6); "_banner.jpg"
            Print #1, "  teaser: /assets/images/RAI_toolkit/"; Left(cellType, 6); ".png"
            Print #1, "sidebar:"
            Print #1, "  - image: /assets/images/RAI_toolkit/"; Left(cellType, 6); ".png"
            Print #1, "    image_alt: """; cellTitle; """"
            'Print #1, "  - title: "" Technology: """
            'Print #1, "    text: "; Range("D" & rowNum).Value; " | "; Range("E" & rowNum).Value; ""
            'Print #1, "  - title: ""Topics: "" "
            'Print #1, "    text: "; Range("F" & rowNum).Value; " | "; Range("G" & rowNum).Value; ""
            'Print #1, "  - title: ""Applications: "" "
            'Print #1, "    text: "; Range("H" & rowNum).Value; " | "; Range("I" & rowNum).Value; ""
            Print #1, "tags:"
            Print #1, "  - "; Range("D" & rowNum).Value
            Print #1, "  - "; Range("E" & rowNum).Value
            Print #1, "  - "; Range("F" & rowNum).Value
            Print #1, "  - "; Range("G" & rowNum).Value
            Print #1, "  - "; Range("H" & rowNum).Value
            Print #1, "  - "; Range("I" & rowNum).Value
            Print #1, "categories:"
            Print #1, "  - "; cellCat
            Print #1, "  - "; cellType
            Print #1, "---"

            'Title
            Print #1, cellDescription
            Print #1,
            Print #1, "[Link]("; Range("N" & rowNum).Value; ")"
            Print #1,
            Print #1, "Source: ["; Range("P" & rowNum).Value; "]("; Range("Q" & rowNum).Value; ")"
            Print #1,
            Print #1, "Ethical Principles: "; Range("J" & rowNum).Value; " | "; Range("K" & rowNum).Value; ""
            Print #1,
            Print #1, "SDGs: "; Range("L" & rowNum).Value; " | "; Range("M" & rowNum).Value; ""
        Close #1
        
        
    Next rowNum
    MsgBox "Rows exported to individual text files."
     
End Sub
