## Overview
Parse the inventory report from Wilco and derive some insights.


## Final Output
Create an excel sheet with the following columns

* SKU
* Description
* OnHand
* Type
* Category
* wt_pc_in_gms
* pcs_per_case
* price_per_pc
* price_per_case
* total_value

## Input
The following files are needed

* Inventory report from Wilco that is sent every friday

* meta/lb_inventory_categories.xlsx.

    * This maps the SKU to the Type and Category.

* private/'PRICE LIST FOR DRY GOODS.xlsx' This is the unmodified from L&B
    * Provides wt/pcs in gms and other pricing info

* private/'PRICE LIST FOR FROZEN GOODS.xlsx'. This is the unmodified ver from L&B
    * Provides wt/pcs in gms and pricing info

## Intermediate Steps

### Clean up  Wilco Inventory
Wilco weekly reports has lot of extra formatiing that needs to be cleaned up. Run the notebook **mk_wilco_inventory_report.ipynb** to create the output inventory list by default *lb_sku_inventory.xlsx*

This list has the following fields

* SKU

* Description

* OnHand

* Type

* Category

### Extract Pricing Info
Clean up LB provided cost sheet to extract the pricing info per SKU. For this run **mk_lb_inventory_pricing.ipynb* to read in the pricing and generate the output *lb_sku_pricing.xlsx*.

This list has the following fields

* SKU
* Description
* wt_pcs_in_case
* pcs_per_case
* price_per_pc

## Final Step










