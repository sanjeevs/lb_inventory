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


