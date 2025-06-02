// Configuration settings
const CONFIG = {
    // Email settings
    DEFAULT_VENDOR_EMAIL: "vendor@example.com",
    YOUR_COMPANY_NAME: "Norton Interiors",
    SENDER_ALIAS_EMAIL: "orders@nortoninteriors.com",
    
    // Sheet settings
    SHEET_NAME: "Pricing",
    HEADER_ROWS: 1,
    
    // Column names
    VENDOR_COL_NAME: "Vendor",
    EMAIL_COL_NAME: "Email",
    PROPERTY_COL_NAME: "Property",
    ROOM_COL_NAME: "Room",
    ITEM_COL_NAME: "Item Name",
    TYPE_COL_NAME: "Item Type",
    QUANTITY_COL_NAME: "Quantity",
    DESCRIPTION_COL_NAME: "Item Name",
    MANUFACTURER_COL_NAME: "Manufacturer",
    SKU_NUMBER_COL_NAME: "SKU",
    DIMENSIONS_COL_NAME: "Dimensions",
    CHECKBOX_COL_NAME: "Request Price",
    STATUS_COL_NAME: "Request Status",
    
    
    // Email settings
    EMAIL_SUBJECT_PREFIX: "Price Request",
    
    // Validation settings
    MAX_ITEMS_PER_EMAIL: 50,
    EMAIL_SEND_TIMEOUT_MS: 30000, // 30 seconds
  };