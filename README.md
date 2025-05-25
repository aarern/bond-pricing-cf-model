# Bond Cash Flow Pricing Model

Description:
This guide helps you dynamically model the cash flows of a fixed-income security and calculate common return metrics (NPV, MOIC, IRR).

1. Enabling VBA Macros:
Open Excel settings (Alt + F T).
Go to the "Customize Ribbon" tab and check the "Developer" box.
Open the "Trust Center" from the Excel settings.
Select "Trust Center Settings."
Go to the "Trusted Locations" tab and select "Add a new location."
Browse to the location where the macro-enabled workbook will be saved.
Enable subfolders of this location to be trusted.
If using a Linux terminal, on the "Trusted Locations" page, check the box that says "Allow Trusted Locations on my network (not recommended)"

2. Naming Excel Sheets, Named Ranges, and Formulas:
Creating Worksheets:
Create four worksheets:
Pricing: Contains bond assumptions and an ActiveX button to execute the script.
Purchase_Schedule: Reflects the initial cash outflow.
Payment_Schedule: Reflects coupon payments and principal payments over the investment’s life.
Cashflow_Model: Stacks arrays from Purchase_Schedule and Payment_Schedule.
Setting Up the Pricing Worksheet:
Add the following descriptors to your rows:

"Today, Maturity Date, Payment Type, Day of Month, Payment Start, Payment Date, Face Value, Coupon Rate, Yield, Input Coupon, Input Number Payments, Calculated Coupon, Calculated Payments, Price, NPV, MOIC, IRR."
Define the named ranges and input formulas as follows:

Today: =TODAY()
Maturity Date: =TODAY() + (365 * 10)
Payment Type: Create a list with values: "Monthly", "Quarterly", "Semi Annual", "Annually".
Day of Month Payment: Create a list with values: "1st", "15th".
Start Payment Date:
excel
Copy code
=IF(payment_date="15th",
  IFS(
    payment_type="Monthly", DATE(YEAR(TODAY()), MONTH(TODAY()) + 1, 15),
    payment_type="Quarterly", DATE(YEAR(TODAY()), MONTH(TODAY()) + 3 - MOD(MONTH(TODAY()) - 1, 3), 15),
    payment_type="Semi Annual", DATE(YEAR(TODAY()), MONTH(TODAY()) + 6 - MOD(MONTH(TODAY()) - 1, 6), 15),
    payment_type="Annually", DATE(YEAR(TODAY()) + 1, MONTH(TODAY()), 15)
  ),
  IFS(
    payment_type="Monthly", DATE(YEAR(TODAY()), MONTH(TODAY()) + 1, 1),
    payment_type="Quarterly", DATE(YEAR(TODAY()), MONTH(TODAY()) + 3 - MOD(MONTH(TODAY()) - 1, 3), 1),
    payment_type="Semi Annual", DATE(YEAR(TODAY()), MONTH(TODAY()) + 6 - MOD(MONTH(TODAY()) - 1, 6), 1),
    payment_type="Annually", DATE(YEAR(TODAY()) + 1, MONTH(TODAY()), 1)
  )
)
Face Value: Custom value input.
Coupon Rate: Custom value input.
Yield: Custom value input.
Input Coupon: Custom value input.
Input Number Payments: Custom value input.
Calculated Coupon: =face_value * coupon_rate
Calculated Payments:
excel
Copy code
=ROUND(
  IFS(
    payment_type="Annually", (maturity_date - today) / 365,
    payment_type="Semi Annual", (maturity_date - today) / (365 / 2),
    payment_type="Quarterly", (maturity_date - today) / (365 / 4),
    payment_type="Monthly", (maturity_date - today) / (365 / 12)
  ),
  0
)
Price:
excel
Copy code
=IFS(
  AND(NOT(ISBLANK(maturity_date)), NOT(ISBLANK(coupon_rate))),
  ((coupon_amount * ((1 - (1 + yield) ^ (-calculated_num_pmt)) / yield)) + (face_value / (1 + yield) ^ calculated_num_pmt)),
  AND(ISBLANK(maturity_date), NOT(ISBLANK(coupon_rate))),
  ((coupon_amount * ((1 - (1 + yield) ^ (-input_num_pmt)) / yield)) + (face_value / (1 + yield) ^ input_num_pmt)),
  AND(NOT(ISBLANK(maturity_date)), ISBLANK(coupon_rate)),
  ((input_coupon * ((1 - (1 + yield) ^ (-calculated_num_pmt)) / yield)) + (face_value / (1 + yield) ^ calculated_num_pmt)),
  AND(ISBLANK(maturity_date), ISBLANK(coupon_rate)),
  ((input_coupon * ((1 - (1 + yield) ^ (-input_num_pmt)) / yield)) + (face_value / (1 + yield) ^ input_num_pmt))
)
NPV: =XNPV(yield, Cashflow_Model!H5#, Cashflow_Model!C5#)
MOIC: =SUMIFS(Cashflow_Model!H5#, Cashflow_Model!H5#, ">0") / ABS(SUMIFS(Cashflow_Model!H5#, Cashflow_Model!H5#, "<0"))
IRR: =XIRR(Cashflow_Model!H5#, Cashflow_Model!C5#)
Setting Up Purchase_Schedule:
In cell C4, create the named range start_purchase.
Populate the cells as follows:
C4: "Period"
D4: "Purchase"
E4: "Coupon Payment"
F4: "Principal Payment"
G4: "Total Payment"
H4: "Cashflow"
I4: "Transaction_Type"
Setting Up Payment_Schedule:
In cell C4, create the named range start_pmt.
Populate the cells as follows:
C4: "Period"
D4: "Purchase"
E4: "Coupon Payment"
F4: "Principal Payment"
G4: "Total Payment"
H4: "Cashflow"
I4: "Transaction_Type"
Setting Up Cashflow_Model:
In cell C4, create the named range start_cf.
Populate the cells as follows:
C4: "Period"
D4: "Purchase"
E4: "Coupon Payment"
F4: "Principal Payment"
G4: "Total Payment"
In cell C5, enter the array formula:
excel
Copy code
=VSTACK(FILTER(Purchase_Schedule!C$5:C$1048576, Purchase_Schedule!C$5:C$1048576<>""), FILTER(Payment_Schedule!C$5:C$1048576, Payment_Schedule!C$5:C$1048576<>""))
Paste formulas for the remaining headers as needed.

3. Setting Up the VBA Code:
Open the VBA environment (Alt + F11).
Click on Sheet1 (Pricing).
Paste the code from "Bond-Pricing-VBA.cls."
[Optional] Insert an ActiveX button on the Pricing sheet. Right-click the button to open the VBA editor, and for the click event, type:
vba
Copy code
Call BondPricing
