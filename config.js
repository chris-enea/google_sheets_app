// Configuration settings
const CONFIG = {
    // Email settings
    DEFAULT_VENDOR_EMAIL: "vendor@example.com",
    YOUR_COMPANY_NAME: "Norton Interiors",
    
    // Sheet settings
    SHEET_NAME: "Sourcing",
    HEADER_ROWS: 1,
    
    // Column names
    VENDOR_COL_NAME: "Vendor",
    EMAIL_COL_NAME: "Email",
    PROPERTY_COL_NAME: "Property",
    ROOM_COL_NAME: "Room",
    ITEM_COL_NAME: "Item",
    TYPE_COL_NAME: "Type",
    QUANTITY_COL_NAME: "Quantity",
    DESCRIPTION_COL_NAME: "Description",
    MANUFACTURER_COL_NAME: "Manufacturer",
    SKU_NUMBER_COL_NAME: "SKU",
    DIMENSIONS_COL_NAME: "Dimensions",
    
    // Column indices (1-based)
    CHECKBOX_COL_INDEX: 1,
    DESCRIPTION_COL_INDEX: 3,
    SKU_NUMBER_COL_INDEX: 6,
    MANUFACTURER_COL_INDEX: 4,
    DIMENSIONS_COL_INDEX: 7,
    STATUS_COL_INDEX: 9,
    
    // Email settings
    EMAIL_SUBJECT_PREFIX: "Price Request",
    
    // Validation settings
    MAX_ITEMS_PER_EMAIL: 50,
    EMAIL_SEND_TIMEOUT_MS: 30000, // 30 seconds
  };