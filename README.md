# PROBLEM STATEMENT

Column A contains the Batch ID, Column B contains the production date, and Column C contains the ship date, as shown here:

![Screenshot 2024-07-08 140555](https://github.com/dannieRope/-Highlight-Rows-based-on-selected-Batch-ID-VBA-Project/assets/132214828/a8b41428-c848-46e3-a9af-7d97a55cc10d)

The Batch ID has a two-digit code to the left of the hyphen and a 3- or 4-digit code to the right of the hyphen.  The first letter of the Batch ID is known as the Identifier and the leading number of the 3- or 4-digit code to the right of the hyphen is known as the Key.  For example, in the Batch ID "N9-363B", the Identifier is "N" and the Key is 3:

![batcjod](https://github.com/dannieRope/-Highlight-Rows-based-on-selected-Batch-ID-VBA-Project/assets/132214828/be3241f0-540b-4d58-9f48-833ac17031be)

The goal is to create a subroutine that allows the user to select the Identifier from a drop-down menu in cell F2 and the Key from a drop-down menu in cell F3  and any rows of the data (columns A, B, and C) whose Batch ID meets those criteria will be highlighted GREEN.

For example, if we start with the worksheet layout above and run the subroutine with the Identifier set to "C" and the Key set to 2, we would get the following:

![Screenshot 2024-07-08 143232](https://github.com/dannieRope/-Highlight-Rows-based-on-selected-Batch-ID-VBA-Project/assets/132214828/575cfb9a-431d-4ca2-819e-32e1e2602818)

## Additional Requirements
- When a match is found in a row, all three columns of data (columns A, B, and C) must be highlighted green.

- When additional rows are added to the data (for example, in row 27 and beyond), your sub should automatically detect the size of data (number of rows) and adjust accordingly.

# SOLUTION 

- Step 1: Creating a drop down menu in Cell F2 for the Identifiers and a drop down in F3 for the keys
  
  To create a drop down, extract distinct identifiers and keys and place them in cell S1 and T1 respectively
  
  Use the following formula to get the unique identifiers and Keys. 
  
  ```
  Identifier in Cell S1
  
  =SORT(UNIQUE(LEFT(A2:A26,1)))
  ```
  ```
  Key in Cell T1
  
  =SORT(UNIQUE(MID(A2:A26,FIND("-",A2:A26)+1,1)))
  ```
 Create the drop down menu in F2 and F3 using Data validation feature available in the data tab in Excel as shown below

 ![Screenshot 2024-07-08 150241](https://github.com/dannieRope/-Highlight-Rows-based-on-selected-Batch-ID-VBA-Project/assets/132214828/9d2e2171-baf6-4f7f-95d2-0aeb9b0b82c5)

 ![Screenshot (20)](https://github.com/dannieRope/-Highlight-Rows-based-on-selected-Batch-ID-VBA-Project/assets/132214828/f2b0c6bf-a5fd-4cd9-972e-622692b6b00d)

- Step 2: Creating a subroutine that highlights rows based on the selected identifier and key

```vba
  Sub HighlightRows()

  Dim TotalRow As Integer
  Dim RowNumber As Integer

 'Getting total number of rows with data excluding the header
    TotalRow = WorksheetFunction.CountA(Columns("A:A")) - 1
    
 'Loop through the rows that satisfy the conditions and highlight rows green
    For RowNumber = 1 To TotalRow
        If Range("Identifier").Value = Identifier(Range("A" & RowNumber + 1)) _
            And Key(Range("A" & RowNumber + 1)) = Range("Key") Then
                Range("A" & RowNumber + 1).Interior.ColorIndex = 4
                Range("A" & RowNumber + 1).Offset(0, 1).Interior.ColorIndex = 4
                Range("A" & RowNumber + 1).Offset(0, 2).Interior.ColorIndex = 4
                  
        End If
        
    Next RowNumber
    
  Range("A1").Select

  End Sub
  ```

The below functions were created to help find both the identifiers and keys. Both are referenced in the subroutine above. 

```vba

Function Identifier(ID As String) As String
Identifier = Left(ID, 1)
End Function
```

```vba
Function Key(ID As String) As Integer
Key = Left(Mid(ID, 4, 4), 1)
End Function
```

- step 3: Creating a subroutine that helps reset or clear all formatings and brings the data to its orginal form.

```vba
Sub Reset()
With Range("A1").CurrentRegion.Cells.Interior
    .Pattern = xlNone
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With
End Sub
```
- Step 4: Creating command button run and reset button
  To insert command button, go to the developer tab, click on insert and choose command button from the drop down as shown below.

  ![Screenshot (21)](https://github.com/dannieRope/-Highlight-Rows-based-on-selected-Batch-ID-VBA-Project/assets/132214828/75b44533-5b0f-4432-9aa6-a086723891f9)

  Name the buttons as run and reset

  Assign the subroutines(micros) to the command buttoons as shown below.

  ![Screenshot (22)](https://github.com/dannieRope/-Highlight-Rows-based-on-selected-Batch-ID-VBA-Project/assets/132214828/d8c99b17-5cd7-43fd-92ac-5af99a118e8e)

  ![Screenshot 2024-07-08 160109](https://github.com/dannieRope/-Highlight-Rows-based-on-selected-Batch-ID-VBA-Project/assets/132214828/275d0cca-fd5d-4d19-b352-a80698ec8f06)

  # CONCLUSION
  
Implimenting the above steps should help achieve exactly what was stated in the problem statement. 

Thanks for reading. 
Find the VBA script (here)[] 





  






