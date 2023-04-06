' © Copyright Iress Limited. All Rights Reserved.
'
' Iress Limited disclaims liability (including for negligence) for loss arising through access or use of this
' product or any information contained in it. Iress Limited does not make any warranties of any kind in
' respect of the product or any information contained in it.
'
' Summary:
' This VB Script example demonstrates retrieving price quote information for each security code specified.
'
' Author: Simon Peckitt
'
' Change History:
' 25-Feb-2014 [Simon Peckitt] Version 1.00: Released.
' 20-Aug-2019 [Ruth Hawkins] Version 1.01: Iress rebrand.

' Dependencies:
' This sample requires Iress Pro Neo 1.09 or higher.

Option Explicit

WScript.Echo "1. Creating the IressServerApi.RequestManager object."
WScript.Echo " "

Dim rm
Set rm = CreateObject("IressServerApi.RequestManager")

WScript.Echo "2. Creating a request object for the PricingQuoteGet method."
WScript.Echo " "

Dim requester
Set requester = rm.CreateMethod("IRESS", "", "PricingQuoteGet", 0)

WScript.Echo "3. Setting input parameters on the PricingQuoteGet method."
WScript.Echo " "

requester.Input.Parameters.Set "SecurityCode", "TEL", 0
requester.Input.Parameters.Set "Exchange", "NZ", 0
requester.Input.Parameters.Set "SecurityCode", "CBA", 1
requester.Input.Parameters.Set "Exchange", "ASX", 1
requester.Input.Parameters.Set "SecurityCode", "IRE", 2
requester.Input.Parameters.Set "Exchange", "ASX", 2
             
WScript.Echo "4. Executing the request."
WScript.Echo " "

requester.Execute

WScript.Echo "5. Getting a list of all available output columns on the PricingQuoteGet method."
WScript.Echo " "

Dim availableFields
availableFields = requester.Output.DataRows.GetAvailableFields()

WScript.Echo "6. Access returned data for all available output columns and all rows."
WScript.Echo " "

Dim data
data = requester.Output.DataRows.GetRows(availableFields, 0, -1)
        
WScript.Echo "7. Output returned data:"
WScript.Echo " "       
 
Dim row
For row = 0 To UBound(data, 1) - LBound(data, 1)
    WScript.Echo " "
    Dim column
    For column = 0 To UBound(availableFields, 1) - LBound(availableFields, 1)
        WScript.Echo chr(9) & availableFields(column) & " [" & data(row, column) & "]"
    Next
Next

WScript.Echo " "
WScript.Echo "8. Done."
WScript.Quit 0
