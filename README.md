# Wayback-Machine-UPC-lookup
This script is designed to look up product names on Internet Archives and then automatically fill in nutritional data when it finds a match from before 2016

To apply this python script to different products, for example ice cream, change some generic word entries in the GENERIC WORDS section of the script from things like "yogurt" to things like "icecream"

Running the code on an excel file with UPCs should produce a new excel file with all matches filled in and a confidence score of high if both the UPC and nutritional information was found on a capture, medium if one or the other was found on a capture, or low if neither of the exact UPC nor the complete nutritional information was found. 
