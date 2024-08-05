VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Click_Update_Models_Click()
    Call BondPricing
End Sub

Sub BondPricing()
    Dim wsPurchaseSchedule As Worksheet
    Dim wsPaymentSchedule As Worksheet
    Dim wsPricing As Worksheet
    Dim wsCashflowModel As Worksheet
    Dim startPurchase As Range
    Dim startPmt As Range
    Dim startCF As Range
    Dim periodDate As Date
    Dim couponRate As Double
    Dim faceValue As Double
    Dim numPmt As Variant
    Dim startPayment As Date
    Dim yield As Double
    Dim todayDate As Date
    Dim paymentType As String
    Dim i As Integer
    Dim cashFlow As Double
    Dim cell As Range
    Dim lastCol As Long
    Dim c As Long
    Dim payment_array As Range
    Dim couponAmount As Variant
    Dim calculatedNumPmt As Variant
    Dim inputNumPmt As Variant
    Dim decimalPart As Double
    Dim lastRow As Long
    Dim firstRow As Long
    Dim firstCol As Long
    Dim inputCoupon As Variant ' Ensure this variable is declared

    ' Error handling
    On Error GoTo ErrorHandler

    ' Set references to the worksheets
    Set wsPricing = ThisWorkbook.Worksheets("Pricing")
    Set wsPurchaseSchedule = ThisWorkbook.Worksheets("Purchase_Schedule")
    Set wsPaymentSchedule = ThisWorkbook.Worksheets("Payment_Schedule")
    Set wsCashflowModel = ThisWorkbook.Worksheets("Cashflow_Model")

    ' Retrieve named ranges from the Pricing worksheet
    With wsPricing
        On Error Resume Next
        couponRate = .Range("coupon_rate").Value
        faceValue = .Range("face_value").Value
        numPmt = .Range("num_pmt").Value
        startPayment = .Range("start_payment").Value
        paymentType = Trim(.Range("payment_type").Value)
        yield = .Range("yield").Value
        todayDate = .Range("today").Value
        On Error GoTo ErrorHandler

        ' Check if named ranges are valid
        If IsError(couponRate) Or IsError(faceValue) Or IsError(numPmt) Or IsError(startPayment) _
            Or IsError(paymentType) Or IsError(yield) Or IsError(todayDate) Then
            MsgBox "One or more of the named ranges on the 'Pricing' worksheet are invalid or missing.", vbCritical
            Exit Sub
        End If
    End With

    ' Retrieve and check named ranges coupon_amount, input_coupon, calculated_num_pmt, and input_num_pmt
    On Error Resume Next
    couponAmount = wsPricing.Range("coupon_amount").Value
    inputCoupon = wsPricing.Range("input_coupon").Value
    calculatedNumPmt = wsPricing.Range("calculated_num_pmt").Value
    inputNumPmt = wsPricing.Range("input_num_pmt").Value
    On Error GoTo ErrorHandler

    If IsError(couponAmount) Or IsError(inputCoupon) Or IsError(calculatedNumPmt) Or IsError(inputNumPmt) Then
        MsgBox "One or more of the named ranges 'coupon_amount', 'input_coupon', 'calculated_num_pmt' or 'input_num_pmt' are invalid or missing.", vbCritical
        Exit Sub
    End If

    ' Use coupon_amount or input_coupon
    If IsEmpty(couponAmount) Then
        couponAmount = inputCoupon
    End If

    If IsEmpty(calculatedNumPmt) Or IsEmpty(inputNumPmt) Then
        MsgBox "One or both of the named ranges 'calculated_num_pmt' or 'input_num_pmt' are empty. Please provide values.", vbExclamation
        Exit Sub
    End If

    ' 1. Populate Purchase Schedule
    On Error Resume Next
    Set startPurchase = wsPurchaseSchedule.Range("start_purchase")
    On Error GoTo ErrorHandler

    If Not startPurchase Is Nothing Then
        ' Clear existing data below start_purchase
        wsPurchaseSchedule.Rows(startPurchase.Row + 1 & ":" & wsPurchaseSchedule.Rows.count).ClearContents

        ' Populate headers in Purchase_Schedule
        startPurchase.Cells(1, 1).Value = "Period"
        startPurchase.Cells(1, 2).Value = "Purchase"
        startPurchase.Cells(1, 3).Value = "Coupon Payment"
        startPurchase.Cells(1, 4).Value = "Principal Payment"
        startPurchase.Cells(1, 5).Value = "Total Payment"
        startPurchase.Cells(1, 6).Value = "Cashflow"
        startPurchase.Cells(1, 7).Value = "Transaction_Type"

        ' Initialize the start date
        periodDate = startPayment

        ' Populate the purchase of the bond in the row below start_purchase
        With startPurchase.Offset(1, 0)
            .Cells(1, 1).Value = periodDate
            .Cells(1, 2).Value = faceValue
            .Cells(1, 3).Value = 0
            .Cells(1, 4).Value = 0
            .Cells(1, 5).Value = faceValue
            .Cells(1, 6).Value = -faceValue
            .Cells(1, 7).Value = "Bond Purchase"
        End With

        ' Format the Period column as dates
        startPurchase.Columns(1).NumberFormat = "mm/dd/yyyy"

        ' Set column width for each populated column
        For c = 1 To 7
            startPurchase.Columns(c).ColumnWidth = 20
        Next c

        ' Determine the last column of data in the row where start_purchase is located
        lastCol = wsPurchaseSchedule.Cells(startPurchase.Row, wsPurchaseSchedule.Columns.count).End(xlToLeft).Column

        ' Apply bottom thick border to all non-blank cells in the start_purchase row
        For c = startPurchase.Column To lastCol
            Set cell = wsPurchaseSchedule.Cells(startPurchase.Row, c)
            If Not IsEmpty(cell.Value) Then
                With cell.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThick
                End With
            End If
        Next c
    Else
        MsgBox "Named range 'start_purchase' is not defined or invalid.", vbCritical
        Exit Sub
    End If

    ' 2. Populate Payment Schedule
    On Error Resume Next
    Set startPmt = wsPaymentSchedule.Range("start_pmt")
    On Error GoTo ErrorHandler

    If Not startPmt Is Nothing Then
        ' Clear existing data in the range where cash flows will be populated
        startPmt.CurrentRegion.ClearContents

        ' Populate headers in Payment_Schedule
        startPmt.Cells(1, 1).Value = "Period"
        startPmt.Cells(1, 2).Value = "Purchase"
        startPmt.Cells(1, 3).Value = "Coupon Payment"
        startPmt.Cells(1, 4).Value = "Principal Payment"
        startPmt.Cells(1, 5).Value = "Total Payment"
        startPmt.Cells(1, 6).Value = "Cashflow"
        startPmt.Cells(1, 7).Value = "Transaction_Type"

        ' Initialize the start date
        periodDate = startPayment

        ' Initialize the row index for coupon payments
        Dim couponStartRow As Long
        couponStartRow = 2

        ' Calculate and populate cash flows for each period
        For i = 1 To calculatedNumPmt
            ' Determine coupon payment
            If i = calculatedNumPmt Then
                ' Final payment includes face value and coupon payment
                cashFlow = couponAmount + faceValue
            Else
                cashFlow = couponAmount
            End If
            
            ' Populate data
            startPmt.Cells(couponStartRow + i - 1, 1).Value = periodDate
            startPmt.Cells(couponStartRow + i - 1, 2).Value = 0
            startPmt.Cells(couponStartRow + i - 1, 3).Value = couponAmount
            
            If i = calculatedNumPmt Then
                ' Final payment includes face value and coupon payment
                startPmt.Cells(couponStartRow + i - 1, 4).Value = faceValue
                startPmt.Cells(couponStartRow + i - 1, 5).Value = cashFlow
                startPmt.Cells(couponStartRow + i - 1, 7).Value = "Coupon + Principal"
                startPmt.Cells(couponStartRow + i - 1, 6).Value = cashFlow
            Else
                startPmt.Cells(couponStartRow + i - 1, 4).Value = 0
                startPmt.Cells(couponStartRow + i - 1, 5).Value = couponAmount
                startPmt.Cells(couponStartRow + i - 1, 7).Value = "Coupon Payment"
                startPmt.Cells(couponStartRow + i - 1, 6).Value = couponAmount
            End If

            ' Increment the period date based on payment type
            Select Case LCase(paymentType)
                Case "monthly"
                    periodDate = DateAdd("m", 1, periodDate)
                Case "quarterly"
                    periodDate = DateAdd("m", 3, periodDate)
                Case "semi annual"
                    periodDate = DateAdd("m", 6, periodDate)
                Case "annually"
                    periodDate = DateAdd("yyyy", 1, periodDate)
                Case Else
                    MsgBox "Unknown payment type: " & paymentType, vbCritical
                    Exit Sub
            End Select
        Next i

        ' Format the Period column as dates
        startPmt.Columns(1).NumberFormat = "mm/dd/yyyy"

        ' Format the output range to display two decimal places
        startPmt.Columns(3).NumberFormat = "0.##"
        startPmt.Columns(4).NumberFormat = "0.##"
        startPmt.Columns(5).NumberFormat = "0.##"
        startPmt.Columns(6).NumberFormat = "0.##"

        ' Set column width for each populated column
        For c = 1 To 7
            startPmt.Columns(c).ColumnWidth = 20
        Next c

        ' Determine the last column with data in the row where start_pmt is located
        lastCol = wsPaymentSchedule.Cells(startPmt.Row, wsPaymentSchedule.Columns.count).End(xlToLeft).Column

        ' Apply bottom thick border to all non-blank cells in the start_pmt row
        For c = startPmt.Column To lastCol
            Set cell = wsPaymentSchedule.Cells(startPmt.Row, c)
            If Not IsEmpty(cell.Value) Then
                With cell.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThick
                End With
            End If
        Next c

    Else
        MsgBox "Named range 'start_pmt' is not defined or invalid.", vbCritical
        Exit Sub
    End If

    ' 3. Apply headers and format columns in Cashflow_Model worksheet
    On Error Resume Next
    Set startCF = wsCashflowModel.Range("start_cf")
    On Error GoTo ErrorHandler

    If Not startCF Is Nothing Then
        ' Determine the last column with data to the right of start_cf
        lastCol = wsCashflowModel.Cells(startCF.Row, wsCashflowModel.Columns.count).End(xlToLeft).Column
        
        ' Apply headers to the start_cf range and neighboring columns
        wsCashflowModel.Cells(startCF.Row, startCF.Column).Value = "Period"
        wsCashflowModel.Cells(startCF.Row, startCF.Column + 1).Value = "Purchase"
        wsCashflowModel.Cells(startCF.Row, startCF.Column + 2).Value = "Coupon Payment"
        wsCashflowModel.Cells(startCF.Row, startCF.Column + 3).Value = "Principal Payment"
        wsCashflowModel.Cells(startCF.Row, startCF.Column + 4).Value = "Total Payment"
        wsCashflowModel.Cells(startCF.Row, startCF.Column + 5).Value = "Cashflow"
        wsCashflowModel.Cells(startCF.Row, startCF.Column + 6).Value = "Transaction_Type"
        
        ' Apply bottom thick border to each cell from start_cf to lastCol
        For c = startCF.Column To lastCol
            Set cell = wsCashflowModel.Cells(startCF.Row, c)
            If Not IsEmpty(cell.Value) Then
                With cell.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThick
                End With
            End If
        Next c
        
        ' Set column width for each populated column
        For c = 0 To 6
            wsCashflowModel.Columns(startCF.Column + c).ColumnWidth = 20
        Next c
    Else
        MsgBox "Named range 'start_cf' is not defined or invalid.", vbCritical
        Exit Sub
    End If

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub


