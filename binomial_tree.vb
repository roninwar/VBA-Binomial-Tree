
'Recursive Function to find the TriangleNumber'
'This is used to find the number of leaves in the tree'  
Public Function TriangleNumb(ByVal iValue As Integer) As Long
   If ((iValue = 0) Or (iValue = 1)) Then
      TriangleNumb = 1
   Else
      TriangleNumb = TriangleNumb(iValue - 1) + iValue
   End If
End Function

'Unoptimized Code to Calculate the value of an Option'
'With up to 250 Steps.' 

'Function Returns an Array of Share Prices Based on the # of Steps'
'And the criterias in the function'
Public Function createSharePTree(steps As Integer, _
val_date As Date, exp_date As Date, sp_0 As Double, _
vol As Double, rf_rate As Double _
 ) As Variant
    
    'Find the total size needed for the Array'
    steps = steps + 1
    
    Dim ArraySize As Integer
    ArraySize = TriangleNumb(steps)
    
    'Initialize the Array'
    Dim arr As Variant
    ReDim arr(ArraySize)
    
    'Place Stock Price at Step 0 in array'
    arr(1) = sp_0
    
    'Initialize Variables for Tree'
    Dim duration As Double
    Dim delta_t As Double
    Dim u_step As Double
    Dim d_step As Double
    
    'Calculation of Variables'
    duration = Application.WorksheetFunction.YearFrac(val_date, exp_date, 0)
    delta_t = duration / (steps - 1)
    u_step = Exp(vol * Sqr(delta_t))
    d_step = 1 / u_step
    

    Dim i As Integer
    Dim j As Integer
    Dim ArrayIndex As Integer
    Dim ArrCalc As Double
    Dim counter As Integer
     
    ArrayIndex = 1
    ArrCalc = 1
    
    For i = 1 To ArraySize
        For j = 1 To i + 1
            If j = 1 Then
                arr(ArrayIndex + j) = arr(ArrCalc) * u_step
            Else
                arr(ArrayIndex + j) = arr(ArrCalc) * d_step
                ArrCalc = ArrCalc + 1
            End If
        Next j
        ArrayIndex = ArrayIndex + i + 1
        If ArrayIndex >= ArraySize Then Exit For
    Next i
  
    createSharePTree = arr
    
End Function

'Function Draws out the a binomial tree with an array'
Public Function drawTree(sheet As String, treeA As Variant, _
 printloc As Integer, Istep As Integer _
)

'Shade the Cells of the Binomial Tree and print out the values'
    
    Dim LeafNumb As Integer
    Dim LeafStart As Integer
    Dim StartStep As Integer
    Dim StepCounter As Integer
    Dim LeafCounter As Integer
    Dim LeafLoc As Integer
    
    Dim i As Integer
   
    ThisWorkbook.Sheets(sheet).Cells(printloc, 1).Font.Bold = True
    ThisWorkbook.Sheets(sheet).Cells(printloc + 1, 1).Font.Bold = True
    ThisWorkbook.Sheets(sheet).Cells(printloc, 1).Value = "Date"
    ThisWorkbook.Sheets(sheet).Cells(printloc + 1, 1).Value = "Steps"
    
    LeafNumb = UBound(treeA, 1) - LBound(treeA, 1)
    steps = Istep + 1
    LeafStart = printloc + 3 + steps - 1
    StartStep = 1 + 1
    StepCounter = 1
    LeafCounter = 0
     
    For i = 1 To LeafNumb
        LeafLoc = LeafStart + LeafCounter
        ThisWorkbook.Sheets(sheet).Cells(LeafLoc, StartStep).Interior.Color = RGB(225, 225, 0)
        ThisWorkbook.Sheets(sheet).Cells(LeafLoc, StartStep).Value = treeA(i)
        StepCounter = StepCounter + 1
        LeafCounter = LeafCounter + 2
        
        If StepCounter = StartStep Then
            StartStep = StartStep + 1
            LeafStart = LeafStart - 1
            LeafCounter = 0
            StepCounter = 1
        End If
    Next i

End Function



