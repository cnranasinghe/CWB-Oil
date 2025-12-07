Public Class CWBConstants

#Region "Database Connection"
    Public Shared CWB_DBCONNECTION_STRING As String = "CWBConnectionString"
#End Region

#Region "Bar Button Hint Constants"
    Public Shared SAVE_HINT As String = "Save records, If any existing records will be over written..."
    Public Shared NEW_HINT As String = "Clears all fields for a new record..."
    Public Shared DELETE_HINT As String = "Delete selected record..."
    Public Shared CLOSE_HINT As String = "Close this window..."
    Public Shared DISPLAYSELECTED_HINT As String = "Display selected record to edit..."
    Public Shared REFRESH_HINT As String = "Refresh current grid..."
    Public Shared PRINT_HINT As String = "Print records on the grid..."




    Public Shared CE_HINT As String = "This Liist Contains Only last Two Months Records. When Other Reocord In Need, Type And Search..."
    Public Shared SEARCH_HINT As String = "Search the Bill..."
    Public Shared UPDATE_HINT As String = "Update existing Bill..."


    Public Shared SHOWRECORD_HINT As String = "Show Records..."



#End Region

#Region "Form Error Hints"
    Public Shared CWB_ERROR_TITLE As String = "Exception Occured"
    Public Shared CWB_ERROR_SAVE_TITLELABEL As String = "Can not save the record"
    Public Shared CWB_ERROR_DELETE_TITLELABEL As String = "Can not delete the record"
    Public Shared CWB_ERROR_DISPLAY_TITLELABEL As String = "Can not display the record"
    Public Shared CWB_ERROR_POPULATE_TITLELABEL As String = "Can not populate records"

#End Region

#Region "Confirmation Hint Constants"
    Public Shared CWB_DELETE_CONFIRMATION_TITLE As String = "Delete the selected record?"
    Public Shared CWB_DELETE_CONFIRMATION_TITLELABEL As String = "It is not recormended to delete the selected record?"
    Public Shared CWB_DELETE_CONFIRMATION_DESCRIPTIONLABEL As String = "Click Yes to Delete"

    Public Shared CWB_UPDATE_CONFIRMATION_TITLE As String = "Record already exists"
    Public Shared CWB_UPDATE_CONFIRMATION_TITLELABEL As String = "Do you want to replace the selected record?"
    Public Shared CWB_UPDATE_CONFIRMATION_DESCRIPTIONLABEL As String = "Click Yes to replace"

    Public Shared CWB_DISPATCH_CONFIRMATION_TITLE As String = "Load No already exists"
    Public Shared CWB_DISPATCH_CONFIRMATION_TITLELABEL As String = "This Load No already sent to factory, Please change the Load No and try again"
    Public Shared CWB_DISPATCH_CONFIRMATION_DESCRIPTIONLABEL As String = "Click OK to continue"

    Public Shared CWB_RECEIVED_CONFIRMATION_TITLE As String = "Check No already exists"
    Public Shared CWB_RECEIVED_CONFIRMATION_TITLELABEL As String = "Check No you are trying to enter is already exists, change the Check No and try again"
    Public Shared CWB_RECEIVED_CONFIRMATION_DESCRIPTIONLABEL As String = "Click OK to continue"

    Public Shared CWB_DELETE_ERROR_CONFIRMATION_TITLE As String = "Can not delete the record"
    Public Shared CWB_DELETE_ERROR_CONFIRMATION_TITLELABEL As String = "This record contains references to other tables, To delete the record, delete the reference records first and try again"
    Public Shared CWB_DELETE_ERROR_CONFIRMATION_DESCRIPTIONLABEL As String = "Click OK to continue"

    Public Shared CWB_SAVESUCCESS_CONFIRMATION_TITLE As String = "Record Saved"
    Public Shared CWB_SAVESUCCESS_CONFIRMATION_TITLELABEL As String = "Record saved succuessfully"
    Public Shared CWB_SAVESUCCESS_CONFIRMATION_DESCRIPTIONLABEL As String = "Click OK to continue"


    Public Shared CWB_PRINT_CONFIRMATION_TITLE As String = "Pring All Records?"
    Public Shared CWB_PRINT_CONFIRMATION_TITLELABEL As String = "You are going to print all the records in the grid?"
    Public Shared CWB_PRINT_CONFIRMATION_DESCRIPTIONLABEL As String = "Click Yes to Print"


#End Region

#Region "Stored Procedures Constatns"
    Public Shared CUSTOMERS_DELETE As String = "Customers_Delete"
    Public Shared CUSTOMERS_GETALL As String = "Customers_GetAll"
    Public Shared CUSTOMERS_INSERT As String = "Customers_Insert"
    Public Shared CUSTOMERS_GETBYID As String = "Customers_GetByID"

    Public Shared EMPLOYERS_DELETE As String = "Employers_Delete"
    Public Shared EMPLOYERS_GETALL As String = "Employers_GetAll"
    Public Shared EMPLOYERS_INSERT As String = "Employers_Insert"
    Public Shared EMPLOYERS_GETBYID As String = "Employers_GetByID"

    Public Shared SUPPLIERS_DELETE As String = "Suppliers_Delete"
    Public Shared SUPPLIERS_GETALL As String = "Suppliers_GetAll"
    Public Shared SUPPLIERS_INSERT As String = "Suppliers_Insert"
    Public Shared SUPPLIERS_GETBYID As String = "Suppliers_GetByID"



    Public Shared STOCK_INSERT As String = "Stock_Insert"
    Public Shared STOCK_DELETE As String = "Stock_Delete"
    Public Shared STOCK_GETALL As String = "Stock_GetAll"
    Public Shared STOCK_GETBYID As String = "Stock_GetByID"

    Public Shared STOCKCATEGORY_GETALL As String = "StockCategory_GetAll"
    Public Shared STOCKCATEGORY_DELETE As String = "StockCategory_Delete"
    Public Shared STOCKCATEGORY_INSERT As String = "StockCategory_Insert"
    Public Shared STOCKCATEGORY_UPDATE As String = "StockCategory_Update"

    Public Shared COMPANY_INSERT As String = "Company_Insert"
    Public Shared COMPANY_GETALL As String = "Company_GetAll"

    Public Shared QUOTATION_INSERT As String = "Quotation_Insert"
    Public Shared QUOTATION_DELETE As String = "Quotation_Delete"
    Public Shared QUOTATION_SELECT_POP As String = "Quotation_Select_Pop"
    Public Shared QUOTATION_SELECTBYID As String = "Quotation_SelectByID"

    Public Shared PURCHASES_INSERT As String = "Purchases_Insert"
    Public Shared PURCHASES_GETALL As String = "Purchases_GetAll"
    Public Shared PURCHASES_GETBYID As String = "Purchases_GetByID"

    Public Shared PURCHASESDESCRIPTION_INSERT As String = "PurchasesDescription_Insert"
    Public Shared PURCHASESDESCRIPTION_GETBYID As String = "PurchasesDescription_GetByID"
    Public Shared PURCHASESDESCRIPTION_DELETE As String = "PurchasesDescription_Delete"

    Public Shared STOCK_REMOVEFROMPURCHASE As String = "Stock_RemoveFromPurchase"
    Public Shared PURCHASES_UPDTESTOCK As String = "Stock_UpdateByPurchase"
    Public Shared PURCHASESCOLLECTION_DELETE As String = "PurchasesCollection_Delete"


    Public Shared PURCHASEORDER_GETALL As String = "PurchaseOrder_GetAll"
    Public Shared PURCHASEORDER_GETBYID As String = "PurchaseOrder_GetByID"
    Public Shared PURCHASEORDER_INSERT As String = "PurchaseOrder_Insert"

    Public Shared PURCHASEORDERDESCRIPTION_DELETE As String = "PurchaseOrderDescription_Delete"
    Public Shared PURCHASEORDERDESCRIPTION_GETBYID As String = "PurchaseOrderDescription_GetByID"
    Public Shared PURCHASEORDERDESCRIPTION_INSERT As String = "PurchaseOrderDescription_Insert"

    Public Shared SALES_INSERT As String = "Sales_Insert"
    Public Shared SALES_GETALL As String = "Sales_GetAll"
    Public Shared SALES_GETBYID As String = "Sales_GetByID"

    Public Shared SALESDESCRIPTION_INSERT As String = "SalesDescription_Insert"
    Public Shared SALESDESCRIPTION_GETBYID As String = "SalesDescription_GetByID"
    Public Shared SALESDESCRIPTION_DELETE As String = "SalesDescription_Delete"

    Public Shared STOCK_UPDATEBYSALES As String = "Stock_UpdateBySales"
    Public Shared STOCK_ADDROMSALES As String = "Stock_AddFromSales"

    Public Shared PURCHASE_UNSETTLED_BILLS As String = "Purchase_Unsettled_Bills"
    Public Shared PURCHASE_SETTLED_BILLS As String = "Purchase_Settled_Bills"

    Public Shared SALES_SETTLED_BILLS As String = "Sales_Settled_Bills"
    Public Shared SALES_UNSETTLED_BILLS As String = "Sales_Unsettled_Bills"


    Public Shared EXPENSES_DELETE As String = "Expenses_Delete"
    Public Shared EXPENSES_GETALL As String = "Expenses_GetAll"
    Public Shared EXPENSES_INSERT As String = "Expenses_Insert"
    Public Shared EXPENSE_GETBYID As String = "Expense_GetByID"


    Public Shared VEHICLES_DELETE As String = "Vehicles_Delete"
    Public Shared VEHICLES_GETALL As String = "Vehicles_GetAll"
    Public Shared VEHICLES_INSERT As String = "Vehicles_Insert"
    Public Shared VEHICLES_GETBYID As String = "Vehicles_GetByID"
    Public Shared VEHICLE_ISEXIST As String = "Vehicle_IsExist"



    Public Shared ACCOUNTS_DELETE As String = "Accounts_Delete"
    Public Shared ACCOUNTS_GETALL As String = "Accounts_GetAll"
    Public Shared ACCOUNTS_GETBYID As String = "Accounts_GetByID"
    Public Shared ACCOUNTS_INSERT As String = "Accounts_Insert"
    Public Shared ACCOUNT_BALANCE As String = "Account_Balance"


    Public Shared PAYMENTS_INSERT As String = "Payments_Insert"
    Public Shared PAYMENTS_GETBYID As String = "Payments_GetByID"
    Public Shared PAYMENTS_GETALL As String = "Payments_GetAll"
    Public Shared PAYMENTS_DELETE As String = "Payments_Delete"
    Public Shared PAYMENTS_GETBYSUPPLIERIDANDDATES As String = "Payments_GetBySupplierIDAndDates"
    Public Shared PAYMENTS_GETBYDATES As String = "Payments_GetByDates"
    Public Shared PAYMENTS_ISCHECKEXISTS As String = "Payments_IsCheckExists"


    Public Shared DISPATCH_ISLOADNOEXISTS As String = "Dispatch_IsLoadNoExists"

    Public Shared RECEIVED_DELETE As String = "Received_Delete"
    Public Shared RECEIVED_GETALL As String = "Received_GetAll"
    Public Shared RECEIVED_GETBYID As String = "Received_GetByID"
    Public Shared RECEIVED_INSERT As String = "Received_Insert"
    Public Shared RECEIVED_ISCHECKEXISTS As String = "Received_IsCheckExists"






    Public Shared SALARY_INSERT As String = "Salary_Insert"
    Public Shared SALARY_DELETE As String = "Salary_Delete"
    Public Shared SALARY_GETALL As String = "Salary_GetAll"
    Public Shared SALARY_GETBYID As String = "Salary_GetByID"
    Public Shared SALARY_GENERATE As String = "Salary_Generate"
    Public Shared SALARY_ISCHECKEXISTS As String = "Salary_IsCheckExists"
    Public Shared SALARY_GETBYEMPLOYERID As String = "Salary_GetByEmployerID"


    Public Shared COLLECTION_INSERT As String = "Collection_Insert"
    Public Shared COLLECTION_DELETE As String = "Collection_Delete"
    Public Shared COLLECTION_GETBYID As String = "Collection_GetByID"
    Public Shared COLLECTION_GETALL As String = "Collection_GetAll"









#End Region

End Class
