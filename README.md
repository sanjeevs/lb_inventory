## Overview
Parse the inventory report from Wilco and derive some insights.

## Installation On Ubuntu
I am using python version 3.8.10. Below are the commands after cloning the repository.

```
python3 -m venv venv

source venv/bin/activate

pip install -r requirements.txt

```

> :warning The **"private"** directory is not part of git hub. This has to be downloaded separately from my google drive.


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


### Extract Pricing Info
Clean up LB provided cost sheet to extract the pricing info per SKU. For this run **mk_lb_inventory_pricing.ipynb* to read in the pricing and generate the output *lb_sku_pricing.xlsx*.

This list has the following fields

* SKU
* Description
* wt_pcs_in_case
* pcs_per_case
* price_per_pc

## Final Step
Run the notebook **mk_lb_inventory_master.ipynb** to read in the intermediate files and write out the fnal master file.

Also reads the following files for categories.

* private/lb_inventory_categories.xlsx.

    * This maps the SKU to the Type and Category.

