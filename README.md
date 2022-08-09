## Overview
Parse the inventory report from Wilco and derive some insights.


## Meta Data
The file "lb_inventory_meta.xlsx" holds all the meta information about the inventory. 

Each item in the inventory has a SKU and is a case. Each case has N pcs. 

It has the following info.
* Location: Warehouse where the item is stored.
* SKU: Each item has a unique SKU.
* Description: item details for humans.
* Type: Whether it is frozen or dry.
* Category: Type of item.
* PcPerCase: Number of pcs n a case.
* PcWtGms: Weight of the pc in grams.
* PricePerPc: Cost of the pc in dollars.
* Expiry: The earliest date at which the item will expire. This is in Month/Date/Year format.
  It is possible that a SKU can have multiple expiry dates since it has different shipments. 
  In such cases this reflects the earliest expiry date.

## Installation On Ubuntu
I am using python version 3.8.10. Below are the commands after cloning the repository.

```
python3 -m venv venv

source venv/bin/activate

pip install -r requirements.txt

```

## Running it
Run the notebook *wilco_stock_report_parser.ipynb* to create a clean version of the stock report.
The file will be created in the out directory.

Then run the notebook *create_master_inventory.ipynb* to create the final output which is 'sample1/out/master_{date}.xlsx'



