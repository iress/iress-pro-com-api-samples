# © Copyright Iress Limited. All Rights Reserved.
#
# Iress Limited disclaims liability (including for negligence) for loss arising through access or
# use of this product or any information contained in it. Iress Limited does not make any warranties
# of any kind in respect of the product or any information contained in it.
#
# Summary:
# This Python Script example demonstrates retrieving price quote information for each security code
# specified. It was built and tested against Python 2.7.
#
# Author: Tim Robinson
#
# Change History:
# 13-Nov-2017 [Tim Robinson] Version 1.00: Released
# 20-Aug-2019 [Ruth Hawkins] Version 1.01: Iress rebrand
#
# Dependencies:
# This sample requires Iress Pro Neo 1.09 or higher.
import win32com.client

print "1. Creating the IressServerApi.RequestManager object."
rqm = win32com.client.Dispatch("IressServerApi.RequestManager")

print "2. Creating a request object for the PricingQuoteGet method."
request = rqm.CreateMethod("IRESS", "", "PricingQuoteGet", 0)
available_fields = request.Output.DataRows.GetAvailableFields()

print "3. Setting input parameters on the PricingQuoteGet method."
request.Input.Header.Set("WaitForResponse", True, 0)
request.Input.Parameters.Set("SecurityCode", "TEL", 0)
request.Input.Parameters.Set("Exchange", "NZ", 0)
request.Input.Parameters.Set("SecurityCode", "IRE", 1)
request.Input.Parameters.Set("Exchange", "ASX", 1)
request.Input.Parameters.Set("SecurityCode", "ABX", 2)
request.Input.Parameters.Set("Exchange", "TSX", 2)

print "4. Execute the request and wait until the data is ready"
# This is a blocking call given the WaitForResponse header is true. Set to false to make async.
# Using a blocking call simplifies the python code.
request.Execute()

print "5. Getting a list of all available output columns on the PricingQuoteGet method."
available_fields = request.Output.DataRows.GetAvailableFields()

print "6. Access returned data for all available output columns and all rows."
row_count = request.Output.DataRows.GetCount()
column_count = len(available_fields)
if row_count > 0:
    data = request.Output.DataRows.GetRows(available_fields, 0, -1)
    print "7. Output returned data"
    for row in range(row_count):
        for column in range(column_count):
            print "\t {}: {}".format(available_fields[column], data[row][column])
        print "\n"
else:
    print "7. Output returned data: No records returned."

print "8. Done."