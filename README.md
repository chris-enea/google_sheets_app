# Norton Project Sheet - Google Apps Script

This Google Apps Script enhances Google Sheets to help manage interior design projects. It provides a user interface for selecting rooms, categories, and items, and then populates a sheet with this information.

## Project Overview

The primary goal of this script is to streamline the creation of project itemization and specification sheets for interior design projects. Users can interact with a custom dialog in Google Sheets to:

*   Define project scope by selecting rooms.
*   Assign category types (e.g., Furniture, Lighting) to rooms.
*   Select specific items within those categories.
*   Specify quantities for each item.
*   Generate a formatted Google Sheet tab with all project items, ready for further detailing and client presentation.

For detailed product requirements, user stories, and features, please see [prd.md](prd.md).

## Scripts Overview

The project is organized into several JavaScript files (`.js`) that manage different aspects of the application:

*   `Code.js`: Main script file, typically containing `onOpen()` and other global functions or initializers.
*   `ui.js`: Handles the creation and management of the user interface (dialogs, sidebars).
*   `itemManager.js`: Contains logic related to managing master lists of items and handling item selections for the current project. (Note: some item-related functions may also be in `sheetManager.js` or `Code.js` depending on specific functionality like master list syncing).
*   `sheetManager.js`: Manages interactions with the Google Sheets, including reading from and writing to various sheets (Master Data, temporary selection sheets, final output sheet).
*   `asana.js`: Contains functions related to Asana integration (if any).
*   `budget.js`: Contains functions related to budget calculations or management (if any).
*   `DashboardScripts.js.html`, `DashboardStyles.css.html`, `folders.js.html`, `modal_scripts.js.html`: These appear to be HTML files used for UI components, likely containing client-side JavaScript and CSS.

**Detailed Function Documentation:** For comprehensive documentation on individual functions, including their purpose, parameters, and return values, please refer to the **JSDoc comments** within each respective `.js` file.

## Setup

Currently, the primary setup involves ensuring the script has the necessary permissions to run and that any required `ScriptProperties` (like the ID of a Master Data Sheet) are correctly configured. (Further details to be added if specific manual setup steps are required).