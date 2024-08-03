# sort_by_picklist
Picklists coming out of ShipStation are ordered by SKU. Labels are not ordered, adding time to pick/pack.\
This script can reorder UPS labels with SKUs to match a picklist's SKU ordering.

## Inputs:
`-p path-to-picklist`\
`-l path-to-labels`\
`-c path-to-conversionfile` a `.xlsx` file with columns `OLD SKU` and `NEW SKU`\
`-o path-to-output`\


## Example Usage
`python sort_labels.py -p /Users/lazer/Desktop/fan_genPick\ List.pdf  -l /Users/lazer/Desktop/fan_genLabels-109744.pdf  -c Conversion\ File.xlsx -o ~/Desktop/output.pdf`

## Troubleshooting
If you have dependency errors, install dependencies with `pip3 install -r requirements.txt`