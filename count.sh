#! /bin/bash

# create a copy of the cyclecount template and rename it for use with python
cp countbckup.xlsx cyclecount.xlsx

# Assign variable for error output when max number of SKUs is reached
pyoutput=$(python3 cyclecount.py)

# Print error if max number of SKU's is reached and script exits
if [ "$pyoutput" == "max-sku-exceeded" ]
then
    echo "Maximum SKU count exceeded. Reduce to 30 or less."
    echo "Script terminated."
else
    echo "Spreadsheet updated."
fi
