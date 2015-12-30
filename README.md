# vba-business-app

# Repo content:

* Main file: test.xlsm. Contains 4 sheets:
         1) generate orders with items
         2) generate business plan from sorted list of orders
         3) generate business plan from random list of orders - 10 samples
         4) GUI - manual assignement for list of random orders
         
* "Orders" folder for storing generated odres

* "Prod_plan" Excel files are the output requested

### VB Code structure
Modules:
    Custom data types (UDT): items and orders
    Utilities for files and data manipulation
    Algorithms for sorting and randomizing
    Plan generator
    Config file for general info: nr of items, machines, orders
    
Forms:
    User form for GUI
Class modules:
    Wrapper for order data type
    Handler for list box events (GUI)
