Sub GenerateSalarySacrificeCalculations()

    ' Declare variables
    Dim EmployeeName As String
    Dim FinancialYear As String
    Dim AnnualSalary As Double
    Dim HasHECS As Boolean
    Dim PayrollCycle As String
    Dim NextPayrollDate As Date
    Dim AmountSacrificed As Double
    Dim PayCyclesRemaining As Integer
    Dim PayCyclesOccurred As Integer
    Dim GrossPay As Double
    Dim NewAnnualIncome As Double
    Dim OriginalTaxableIncome As Double
    Dim NewTaxableIncome As Double
    Dim OriginalIncomeTax As Double
    Dim NewIncomeTax As Double
    Dim OriginalHECSTax As Double
    Dim NewHECSTax As Double
    Dim OriginalMedicareLevy As Double
    Dim NewMedicareLevy As Double
    Dim TaxPaidToDate As Double
    Dim TaxPaidToDateHECS As Double
    Dim TaxPaidToDateMedicare As Double
    Dim RemainingIncomeTax As Double
    Dim RemainingHECSTax As Double
    Dim RemainingMedicareLevy As Double
    Dim RemainingTotalTax As Double
    Dim OriginalNetPay As Double
    Dim NewNetPay As Double
    Dim ws As Worksheet
    Dim comparisonWs As Worksheet
    Dim scheduleWs As Worksheet
    Dim i As Integer
    Dim PayDate As Date
    Dim EndOfFinancialYear As Date
    Dim StartOfFinancialYear As Date
    Dim PayCyclesInYear As Integer

    On Error GoTo ErrorHandler

    ' Set the inputs from the table on the first sheet
    Set ws = ThisWorkbook.Sheets("Donation_Tax_Calc")
    EmployeeName = ws.Range("B8").Value
    FinancialYear = ws.Range("B9").Value
    AnnualSalary = CDbl(ws.Range("B10").Value)
    HasHECS = (LCase(ws.Range("B11").Value) = "yes")
    PayrollCycle = ws.Range("B12").Value
    NextPayrollDate = CDate(ws.Range("B13").Value)
    AmountSacrificed = CDbl(ws.Range("B14").Value)

    ' Set the start and end of the financial year
    StartOfFinancialYear = DateSerial(Year(NextPayrollDate), 7, 1)
    EndOfFinancialYear = DateSerial(Year(NextPayrollDate) + 1, 6, 30)

    ' Validate the NextPayrollDate
    If NextPayrollDate < StartOfFinancialYear Or NextPayrollDate > EndOfFinancialYear Then
        MsgBox "Date outside of FY25", vbExclamation
        Exit Sub
    End If

    ' Determine pay cycles
    If LCase(PayrollCycle) = "fortnightly" Then
        PayCyclesInYear = 26
        PayCyclesOccurred = Int((NextPayrollDate - StartOfFinancialYear) / 14)
        PayCyclesRemaining = PayCyclesInYear - PayCyclesOccurred
        PayCyclePay = AnnualSalary / PayCyclesInYear
    Else ' Assume monthly if not fortnightly
        PayCyclesInYear = 12
        PayCyclesOccurred = DateDiff("m", StartOfFinancialYear, NextPayrollDate)
        PayCyclesRemaining = PayCyclesInYear - PayCyclesOccurred
        PayCyclePay = AnnualSalary / PayCyclesInYear
    End If

    ' Calculate the original and new annual taxable incomes
    OriginalTaxableIncome = AnnualSalary
    NewTaxableIncome = AnnualSalary - (AmountSacrificed * PayCyclesRemaining)

    ' Calculate the original and new yearly tax liabilities
    OriginalIncomeTax = CalculateIncomeTax(OriginalTaxableIncome)
    NewIncomeTax = CalculateIncomeTax(NewTaxableIncome)

    OriginalHECSTax = IIf(HasHECS, CalculateHECS(OriginalTaxableIncome), 0)
    NewHECSTax = IIf(HasHECS, CalculateHECS(NewTaxableIncome), 0)

    OriginalMedicareLevy = CalculateMedicare(OriginalTaxableIncome)
    NewMedicareLevy = CalculateMedicare(NewTaxableIncome)

    ' Calculate the tax paid to date
    TaxPaidToDate = (OriginalIncomeTax / PayCyclesInYear) * PayCyclesOccurred
    TaxPaidToDateHECS = (OriginalHECSTax / PayCyclesInYear) * PayCyclesOccurred
    TaxPaidToDateMedicare = (OriginalMedicareLevy / PayCyclesInYear) * PayCyclesOccurred

    ' Calculate the remaining tax to be paid for the rest of the year
    RemainingIncomeTax = NewIncomeTax - TaxPaidToDate
    RemainingHECSTax = NewHECSTax - TaxPaidToDateHECS
    RemainingMedicareLevy = NewMedicareLevy - TaxPaidToDateMedicare

    ' Calculate total remaining tax per cycle
    RemainingTotalTax = (RemainingIncomeTax + RemainingHECSTax + RemainingMedicareLevy) / PayCyclesRemaining

    ' Calculate the original and new net pay per cycle
    OriginalNetPay = PayCyclePay - (OriginalIncomeTax / PayCyclesInYear + OriginalHECSTax / PayCyclesInYear + OriginalMedicareLevy / PayCyclesInYear)
    NewNetPay = (PayCyclePay - AmountSacrificed) - RemainingTotalTax

    ' Create a new worksheet for the comparison
    Set comparisonWs = ThisWorkbook.Sheets.Add
    comparisonWs.Name = EmployeeName & "-Comparison"

    ' First Table: Annual Summary
    comparisonWs.Range("A1").Value = "Description"
    comparisonWs.Range("B1").Value = "Original"
    comparisonWs.Range("C1").Value = "With Salary Sacrifice"
    comparisonWs.Range("A2").Value = "Gross Pay per Annum"
    comparisonWs.Range("A3").Value = "Total Salary Sacrifice This Year"
    comparisonWs.Range("A4").Value = "Taxable Income"
    comparisonWs.Range("A5").Value = "Total Income Tax for Year"
    comparisonWs.Range("A6").Value = "Total HECS-HELP for Year"
    comparisonWs.Range("A7").Value = "Total Medicare Levy for Year"
    comparisonWs.Range("A8").Value = "Total Tax for Year"

    comparisonWs.Range("B2").Value = OriginalTaxableIncome
    comparisonWs.Range("B3").Value = 0
    comparisonWs.Range("B4").Value = OriginalTaxableIncome
    comparisonWs.Range("B5").Value = OriginalIncomeTax
    comparisonWs.Range("B6").Value = OriginalHECSTax
    comparisonWs.Range("B7").Value = OriginalMedicareLevy
    comparisonWs.Range("B8").Value = OriginalIncomeTax + OriginalHECSTax + OriginalMedicareLevy

    comparisonWs.Range("C2").Value = OriginalTaxableIncome
    comparisonWs.Range("C3").Value = AmountSacrificed * PayCyclesRemaining
    comparisonWs.Range("C4").Value = NewTaxableIncome
    comparisonWs.Range("C5").Value = NewIncomeTax
    comparisonWs.Range("C6").Value = NewHECSTax
    comparisonWs.Range("C7").Value = NewMedicareLevy
    comparisonWs.Range("C8").Value = NewIncomeTax + NewHECSTax + NewMedicareLevy

    ' Second Table: Information to Date
    comparisonWs.Range("A10").Value = "Description"
    comparisonWs.Range("B10").Value = "Information to Date"
    comparisonWs.Range("A11").Value = "Pay Cycles That Have Occurred This Year"
    comparisonWs.Range("A12").Value = "Pay Cycles to Come This Year"
    comparisonWs.Range("A13").Value = "Gross Income Paid to Date"
    comparisonWs.Range("A14").Value = "Income Tax Paid to Date"
    comparisonWs.Range("A15").Value = "HECS-HELP Paid to Date"
    comparisonWs.Range("A16").Value = "Medicare Levy Paid to Date"
    comparisonWs.Range("A17").Value = "Total Tax Paid to Date"

    comparisonWs.Range("B11").Value = PayCyclesOccurred
    comparisonWs.Range("B12").Value = PayCyclesRemaining
    comparisonWs.Range("B13").Value = PayCyclePay * PayCyclesOccurred
    comparisonWs.Range("B14").Value = TaxPaidToDate
    comparisonWs.Range("B15").Value = TaxPaidToDateHECS
    comparisonWs.Range("B16").Value = TaxPaidToDateMedicare
    comparisonWs.Range("B17").Value = TaxPaidToDate + TaxPaidToDateHECS + TaxPaidToDateMedicare

    ' Third Table: Remaining Amounts After Salary Sacrifice
    comparisonWs.Range("A19").Value = "Description"
    comparisonWs.Range("B19").Value = "Remaining Amounts After Salary Sacrifice"
    comparisonWs.Range("A20").Value = "Gross Pay Remaining This Year"
    comparisonWs.Range("A21").Value = "Income Tax Remaining This Year"
    comparisonWs.Range("A22").Value = "HECS-HELP Remaining This Year"
    comparisonWs.Range("A23").Value = "Medicare Levy Remaining This Year"
    comparisonWs.Range("A24").Value = "Total Tax Remaining This Year"

    comparisonWs.Range("B20").Value = PayCyclePay * PayCyclesRemaining
    comparisonWs.Range("B21").Value = RemainingIncomeTax
    comparisonWs.Range("B22").Value = RemainingHECSTax
    comparisonWs.Range("B23").Value = RemainingMedicareLevy
    comparisonWs.Range("B24").Value = RemainingTotalTax * PayCyclesRemaining

    ' Fourth Table: Per Cycle Breakdown
    comparisonWs.Range("A26").Value = "Description"
    comparisonWs.Range("B26").Value = "Original"
    comparisonWs.Range("C26").Value = "With Salary Sacrifice"
    comparisonWs.Range("A27").Value = "Gross Pay per Cycle"
    comparisonWs.Range("A28").Value = "Taxable Income per Cycle"
    comparisonWs.Range("A29").Value = "Income Tax per Cycle"
    comparisonWs.Range("A30").Value = "HECS-HELP per Cycle"
    comparisonWs.Range("A31").Value = "Medicare Levy per Cycle"
    comparisonWs.Range("A32").Value = "Total Tax per Cycle"
    comparisonWs.Range("A33").Value = "Net Pay per Cycle"

    comparisonWs.Range("B27").Value = PayCyclePay
    comparisonWs.Range("B28").Value = OriginalTaxableIncome / PayCyclesInYear
    comparisonWs.Range("B29").Value = OriginalIncomeTax / PayCyclesInYear
    comparisonWs.Range("B30").Value = OriginalHECSTax / PayCyclesInYear
    comparisonWs.Range("B31").Value = OriginalMedicareLevy / PayCyclesInYear
    comparisonWs.Range("B32").Value = OriginalIncomeTax / PayCyclesInYear + OriginalHECSTax / PayCyclesInYear + OriginalMedicareLevy / PayCyclesInYear
    comparisonWs.Range("B33").Value = OriginalNetPay

    comparisonWs.Range("C27").Value = PayCyclePay
    comparisonWs.Range("C28").Value = NewTaxableIncome / PayCyclesInYear
    comparisonWs.Range("C29").Value = RemainingIncomeTax / PayCyclesRemaining
    comparisonWs.Range("C30").Value = RemainingHECSTax / PayCyclesRemaining
    comparisonWs.Range("C31").Value = RemainingMedicareLevy / PayCyclesRemaining
    comparisonWs.Range("C32").Value = RemainingTotalTax
    comparisonWs.Range("C33").Value = NewNetPay

    ' Format the comparison table
    comparisonWs.Range("A1:C33").Columns.AutoFit
    comparisonWs.Range("B2:C33").NumberFormat = "$#,##0.00"
    
    ' Create a new worksheet for the pay cycle schedule
    Set scheduleWs = ThisWorkbook.Sheets.Add
    scheduleWs.Name = EmployeeName & "-Pay Schedule"
    
    ' Set up the schedule table headers
    scheduleWs.Range("A1").Value = "Pay Cycle Date"
    scheduleWs.Range("B1").Value = "Gross Pay"
    scheduleWs.Range("C1").Value = "Amount Sacrificed"
    scheduleWs.Range("D1").Value = "Taxable Income"
    scheduleWs.Range("E1").Value = "Income Tax"
    scheduleWs.Range("F1").Value = "HECS-HELP"
    scheduleWs.Range("G1").Value = "Medicare Levy"
    scheduleWs.Range("H1").Value = "Net Pay"
    
    ' Initialize the pay date
    PayDate = NextPayrollDate
    
    ' Loop through each remaining pay cycle and output the results
    For i = 1 To PayCyclesRemaining
        If PayDate > EndOfFinancialYear Then Exit For
        
        scheduleWs.Cells(i + 1, 1).Value = PayDate
        scheduleWs.Cells(i + 1, 2).Value = PayCyclePay
        scheduleWs.Cells(i + 1, 3).Value = AmountSacrificed
        scheduleWs.Cells(i + 1, 4).Value = PayCyclePay - AmountSacrificed
        scheduleWs.Cells(i + 1, 5).Value = RemainingIncomeTax / PayCyclesRemaining
        scheduleWs.Cells(i + 1, 6).Value = RemainingHECSTax / PayCyclesRemaining
        scheduleWs.Cells(i + 1, 7).Value = RemainingMedicareLevy / PayCyclesRemaining
        scheduleWs.Cells(i + 1, 8).Value = NewNetPay
        
        ' Increment the pay date for the next cycle
        If LCase(PayrollCycle) = "fortnightly" Then
            PayDate = DateAdd("ww", 2, PayDate)
        Else ' Assume monthly if not fortnightly
            PayDate = DateAdd("m", 1, PayDate)
        End If
    Next i
    
    ' Format the schedule table
    scheduleWs.Range("A1:H" & i).Columns.AutoFit
    scheduleWs.Range("B2:H" & i).NumberFormat = "$#,##0.00"
    
    MsgBox "Calculation completed and exported to the worksheets: '" & comparisonWs.Name & "' and '" & scheduleWs.Name & "'."
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical

End Sub

Function CalculateIncomeTax(AnnualIncome As Double) As Double
    ' Income tax brackets (FY25, simplified for illustration)
    Dim TaxRate As Double
    Dim Tax As Double
    
    ' Calculate the Income Tax based on the annual income
    If AnnualIncome <= 18200 Then
        TaxRate = 0
        Tax = 0
    ElseIf AnnualIncome <= 45000 Then
        TaxRate = 0.16
        Tax = (AnnualIncome - 18200) * TaxRate
    ElseIf AnnualIncome <= 135000 Then
        TaxRate = 0.3
        Tax = 4288 + (AnnualIncome - 45000) * TaxRate
    ElseIf AnnualIncome <= 190000 Then
        TaxRate = 0.37
        Tax = 31288 + (AnnualIncome - 135000) * TaxRate
    Else
        TaxRate = 0.45
        Tax = 51638 + (AnnualIncome - 190000) * TaxRate
    End If
    
    CalculateIncomeTax = Tax
End Function

Function CalculateHECS(AnnualIncome As Double) As Double
    ' HECS repayment thresholds and rates for FY25
    Dim HECSRate As Double
    Dim HECS As Double
    
    Select Case AnnualIncome
        Case Is < 54435
            HECSRate = 0
        Case 54435 To 62850
            HECSRate = 0.01
        Case 62851 To 66620
            HECSRate = 0.02
        Case 66621 To 70618
            HECSRate = 0.025
        Case 70619 To 74855
            HECSRate = 0.03
        Case 74856 To 79346
            HECSRate = 0.035
        Case 79347 To 84107
            HECSRate = 0.04
        Case 84108 To 89154
            HECSRate = 0.045
        Case 89155 To 94503
            HECSRate = 0.05
        Case 94504 To 100174
            HECSRate = 0.055
        Case 100175 To 106185
            HECSRate = 0.06
        Case 106186 To 112556
            HECSRate = 0.065
        Case 112557 To 119309
            HECSRate = 0.07
        Case 119310 To 126467
            HECSRate = 0.075
        Case 126468 To 134056
            HECSRate = 0.08
        Case 134057 To 142100
            HECSRate = 0.085
        Case 142101 To 150626
            HECSRate = 0.09
        Case 150627 To 159663
            HECSRate = 0.095
        Case Is >= 159664
            HECSRate = 0.1
        Case Else
            HECSRate = 0 ' Default case (should not be reached)
    End Select
    
    HECS = AnnualIncome * HECSRate
    
    CalculateHECS = HECS
End Function

Function CalculateMedicare(AnnualIncome As Double) As Double
    ' Medicare levy is 2% of taxable income
    CalculateMedicare = AnnualIncome * 0.02
End Function

