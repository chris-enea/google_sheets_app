# PRD: Norton Project Sheet

## 1. Product overview
### 1.1 Document title and version
- PRD: Norton Project Sheet
- Version: 1.4.2 (Added FFE to Pricing sheet data copy)

### 1.2 Product summary
   - This Google App Script assists users in building and maintaining a Google Sheet to manage interior design projects using a dialog UI. It allows users to select rooms for the project and manage those rooms (add, update, and delete).
   - The script enables users to select category types for each room and manage those types (add, update, and delete). It also allows users to select items in each room and manage those items (add, update, and delete).
   - Finally, the tool populates a Google Sheet with the selected items, streamlining the project management process for interior designers.

## 2. Goals
### 2.1 Business goals
   - Increase efficiency in project setup and management.
   - Improve accuracy of project itemization and specification.
   - Enhance client communication by providing clear, comprehensive project sheets.
### 2.2 User goals
   - Quickly generate comprehensive project item lists for interior design projects.
   - Easily modify project scope by adding, updating, or removing rooms, categories, and items.
   - Maintain an organized and up-to-date record of all project elements within a Google Sheet.
### 2.3 Non-goals
   - This version will not provide 3D visualization of rooms.
   - This version will not integrate with supplier inventory systems.
   - Direct budget management and invoicing are not part of the current scope (budget management is a potential future enhancement).
## 3. User personas
### 3.1 Key user types
   - Lead Interior Designer
   - Design Assistant
### 3.2 Basic persona details
   - **Lead Interior Designer**: Responsible for the overall project execution, from concept to completion. Uses the tool to define project scope, select and specify all items, and generate the final project sheet for client and internal use.
   - **Design Assistant**: Supports the Lead Interior Designer in project execution. Uses the tool for data entry, updating item details, managing room and category lists, and generating interim or draft project sheets.
### 3.3 Role-based access
   - **All Users**: Currently, there is no user authentication or role-based access control. All users of the Google Sheet and associated App Script will have the same level of access to create, read, update, and delete project data.
## 4. Functional requirements
   - **Project Initialization & Configuration** (Priority: High)
     - **Master Template Setup (Manual, One-Time Admin Task):**
       - The designated "Master Template" Google Sheet requires specific script properties to be manually set by an administrator via File > Project Properties > Script Properties:
         - `IS_MASTER_TEMPLATE` (String): Set to `'true'`. Identifies this sheet as the master source.
         - `MASTER_TEMPLATE_ACTUAL_ID` (String): Set to the actual Google Sheet file ID of this Master Template itself. Used for self-identification.
         - `DATA_SHEET_ID` (String): Set to the file ID of the shared Master Data Sheet (which contains master lists for rooms, categories, items, etc.). This ID will be inherited by all new projects created from this template.
       - Important: Project-specific properties like `PROJECT_NAME` and `PROJECT_INITIALIZED` must NOT be set or must be cleared from the Master Template's script properties. The `DATA_SHEET_ID` set here is the *default* for new projects.
     - **Project Creation from Master (User-Initiated via UI):**
       - When the Master Template sheet is opened, the `onOpen` script checks if `IS_MASTER_TEMPLATE` is `'true'` and if `MASTER_TEMPLATE_ACTUAL_ID` matches the current file's ID.
       - If true, it presents the user with a choice (e.g., UI prompt or custom menu): "Create NEW project from this template" or "Edit Master Template content".
       - If "Create NEW project" is chosen:
         - The script prompts the user to enter a "Project Name".
         - The script then creates a copy of the active Master Template sheet using `SpreadsheetApp.getActiveSpreadsheet().copy('[Project Name] - Budget Sheet')`.
         - Crucially, this copy inherits all script properties from the Master Template, including `IS_MASTER_TEMPLATE`, `MASTER_TEMPLATE_ACTUAL_ID`, and the pre-configured `DATA_SHEET_ID`.
         - The user is alerted that the new project file has been created and should be opened to complete initialization.
     - **New Project Copy Initialization (Automatic on First Open of the Copy):**
       - When a newly created project copy (e.g., "My Client Project - Budget Sheet") is opened by any user for the first time:
         - The `onOpen` script in the copy detects its state: `IS_MASTER_TEMPLATE` is `'true'` (inherited), `MASTER_TEMPLATE_ACTUAL_ID` is present (inherited) but does *not* match the *current* file's ID, and `PROJECT_INITIALIZED` is not `'true'` (or is absent).
         - The script automatically extracts the base "Project Name" from the current file name (e.g., "My Client Project" from "My Client Project - Budget Sheet").
         - The script **does not prompt the user for the Master Data Sheet ID**. It uses the `DATA_SHEET_ID` value that was inherited directly from the Master Template's script properties.
         - The script then finalizes the project copy's configuration by setting its script properties:
           - `PROJECT_NAME`: Set to the extracted base name (e.g., "My Client Project").
           - `PROJECT_INITIALIZED`: Set to `'true'`. This marks the project as fully configured.
           - `IS_MASTER_TEMPLATE`: Set to `'false'`. This sheet is now a project, not a master.
           - `MASTER_TEMPLATE_ACTUAL_ID`: This property is deleted from the project copy as it's no longer relevant.
           - The `DATA_SHEET_ID` (inherited from master) is retained and is now the active data sheet for this project.
         - An alert message confirms successful initialization, e.g., "Project 'My Client Project' has been initialized using Data Sheet ID: [value of inherited DATA_SHEET_ID]. Standard menus are now active."
     - **Initialized Project Operation (Normal Use):**
       - On subsequent opens of an initialized project (where `PROJECT_INITIALIZED` is `'true'`):
         - The `onOpen` script confirms `IS_MASTER_TEMPLATE` is `false` and `MASTER_TEMPLATE_ACTUAL_ID` is absent (and may perform self-correction if these are unexpectedly found, logging a warning).
         - The script proceeds to load standard project menus and functionalities, using the `PROJECT_NAME` and `DATA_SHEET_ID` stored in its own script properties.
     - **Manual Project Initialization (Fallback for Unconfigured or Externally Copied Sheets):**
       - A menu option, e.g., `Project Manager > Setup > Initialize Sheet as Project Manually`, is available for sheets that are not the Master Template and are not yet initialized (e.g., a sheet copied manually outside the script, or if the automatic initialization failed for some reason).
       - This function will:
         - Prompt for "Project Name", suggesting a default extracted from the current file name if possible.
         - Check if a `DATA_SHEET_ID` property already exists (e.g., inherited from a master if it was a copy that failed to auto-initialize). If found, it may confirm with the user if they want to use this ID, or prompt for a new one if it seems invalid or is missing.
         - If no valid `DATA_SHEET_ID` is found or confirmed, it will prompt the user to input the `DATA_SHEET_ID` manually.
         - Once a valid Project Name and `DATA_SHEET_ID` are obtained, it sets/updates the script properties: `PROJECT_NAME`, `DATA_SHEET_ID`, `PROJECT_INITIALIZED='true'`, `IS_MASTER_TEMPLATE='false'`, and ensures `MASTER_TEMPLATE_ACTUAL_ID` is cleared/deleted.
         - It may also offer to rename the current file to the standard `[Project Name] - Budget Sheet` format if it doesn't already match.
     - **Edit Project Configuration (for Initialized Projects):**
       - An existing menu option, e.g., `Project Manager > Setup > Edit Project Configuration`, allows the user to view and update the `PROJECT_NAME` and `DATA_SHEET_ID` for an already initialized project. This is useful for correcting errors or changing the linked Master Data Sheet post-initialization.
   - **Master Data Management (via UI interacting with Master Data Sheet)** (Priority: High)
     - **Master Rooms List (from Master Data Sheet):**
       - Allow user to add new room names to the master list in the Master Data Sheet (e.g., via `getRooms`, `getRoomNamesFromSheet`).
       - Allow user to edit existing room names in the master list.
       - When a room is deleted via the UI, it is removed from the Master Data Sheet and from the `_TempSelectedRooms` sheet in the Active Project Sheet.
     - **Master Category Types List (from Master Data Sheet):**
       - Allow user to add new category types (e.g., "Furniture," "Lighting") to the master list in the Master Data Sheet (e.g., via `getTypes`).
       - Allow user to edit names of existing category types in the master list.
       - When a category type is deleted via the UI, it is removed from the Master Data Sheet and from the `_TempRoomTypes` sheet in the Active Project Sheet. (Deletion is prevented if the Category Type is currently assigned as an `Item-Type` to any items in the Master Items List).
     - **Master Items List (from Master Data Sheet):**
       - The Master Data Sheet contains an "Items" list/tab with columns: `Item-Type` (referencing a Master Category Type) and `Item-Name`.
       - Allow user to add new items (with their `Item-Type` and `Item-Name`) to the master list in the Master Data Sheet (e.g., via `getAvailableItems`).
       - Allow user to edit the `Item-Type` and `Item-Name` of existing items in the master list.
       - When an item is deleted via the UI, it is removed from the Master Data Sheet and from the `_TempItemSelections` sheet in the Active Project Sheet.
   - **Current Project Building (UI populating temporary sheets in Active Project Sheet)** (Priority: High)
     - **Room Selection for Current Project:**
       - Allow user to select rooms (from the Master Rooms List) to include in the current design project. Selections are stored temporarily in the `_TempSelectedRooms` sheet in the Active Project Sheet.
       - Display the list of selected rooms for the current project.
       - Display a bulleted list of assigned categories under each room in the selected rooms list in the sidebar.
       - Allow user to remove a room from the current project (updates `_TempSelectedRooms`; does not delete from master list).
     - **Category Type Assignment for Current Project Rooms:**
       - For each selected room in the current project, allow user to assign category types (from the Master Category Types List). Selections stored temporarily in the `_TempRoomTypes` sheet.
       - Allow user to remove a category type from a specific room in the current project (updates `_TempRoomTypes`; does not delete from master list).
     - **Item Selection & Configuration for Current Project Rooms/Categories:**
       - For each assigned category in a room of the current project, allow user to select items (referencing `Item-Name` from the Master Items List, filtered by `Item-Type` matching the selected Category Type). Selections stored temporarily in the `_TempItemSelections` sheet.
       - When an item is selected for the current project, allow user to specify its `QUANTITY`. This is the primary piece of data managed by the UI for an item instance in the project.
       - Allow user to edit the `QUANTITY` for an item in the current project (updates `_TempItemSelections`).
       - Allow user to remove an item from a category in the current project (updates `_TempItemSelections`; does not delete from master list).
   - **Sheet Generation/Population (in Active Project Sheet)** (Priority: High)
     - Automatically or on-demand populate a designated output tab in the Active Project Sheet using data from the temporary sheets (`_TempSelectedRooms`, `_TempRoomTypes`, `_TempItemSelections`).
     - The sheet should list items grouped by Room, then by Category (Type).
     - The script will populate the following columns: `ROOM`, `TYPE`, `ITEM`, `QUANTITY`.
     - The script will also create headers for the following columns, which are intended for manual user input directly in the sheet: `LOW`, `LOW TOTAL`, `HIGH`, `HIGH TOTAL`, `SPEC/FFE`.
     - Ensure data in the script-populated columns (`ROOM`, `TYPE`, `ITEM`, `QUANTITY`) updates dynamically or via a "refresh/re-populate" action when changes are made to the current project configuration in the UI (and thus to the temporary sheets).
     - **FFE to Pricing Sheet Data Copy (New):** After the "FFE" sheet is generated, the script will automatically create/update a sheet named "Pricing". It will copy the following columns from the "FFE" sheet: `ROOM`, `TYPE`, `ITEM`, `QUANTITY`, `LOW TOTAL`, `HIGH TOTAL`. These will be copied to the "Pricing" sheet with the headers: `Room`, `Item Type`, `Item Name`, `Quantity`, `Budget Low`, `Budget High` respectively, on a row-by-row basis.
   - **User Interface (UI) Management** (Priority: High)
     - Provide a clear, intuitive, and responsive dialog UI for all management and selection functions.
     - UI should clearly distinguish between managing master data and building the current project.
     - Ensure UI elements are consistently styled and easy to understand.
     - Provide feedback to the user for actions (e.g., "Master Room added," "Item selected for project," "Error saving configuration").
     - **Room Category Assignment in Main Content**: Users can manage room category assignments directly in the main content area. This includes selecting a project room and then assigning/unassigning master category types to it. (New)
## 5. User experience
### 5.1. Entry points & first-time user flow
- Users access the tool via a custom menu item in Google Sheets. The `onOpen` function dynamically determines the sheet's state (Master Template, New Project Copy pending initialization, Initialized Project, or Unconfigured Sheet) and adjusts available menu options and behavior accordingly.
- **Master Template First-Time Setup (Manual Admin Task - One Time):**
  - An administrator or designated setup user manually opens the script editor for the Google Sheet file intended to be the Master Template.
  - Navigates to File > Project Properties > Script Properties tab.
  - Adds (or edits if existing) the following script properties:
    - Key: `IS_MASTER_TEMPLATE`, Value: `true` (as a string: `'true'`)
    - Key: `MASTER_TEMPLATE_ACTUAL_ID`, Value: The actual file ID of this Master Template sheet (copied from its URL)
    - Key: `DATA_SHEET_ID`, Value: The file ID of the shared Master Data Sheet that new projects will use by default
  - Ensures any project-specific properties like `PROJECT_NAME` or `PROJECT_INITIALIZED` are NOT present or are explicitly cleared from the Master Template's script properties.
- **User Creating a New Project (from the Master Template):**
  - User opens the designated Master Template Google Sheet.
  - The `onOpen` script detects it's the master. A UI prompt (or custom menu) appears: "This is the Master Template. What would you like to do?" with options like "Create a NEW project from this template" and "Continue editing Master Template content".
  - User selects "Create a NEW project...".
  - A prompt asks: "Please enter the base Project Name for the new project:".
  - User enters, for example, "Johnson Residence Phase 1".
  - The script executes `SpreadsheetApp.getActiveSpreadsheet().copy('Johnson Residence Phase 1 - Budget Sheet')`.
  - An alert confirms: "New project file 'Johnson Residence Phase 1 - Budget Sheet' has been created. Please open this new file from your Google Drive to complete its setup."
- **User Opening a New Project Copy for the First Time (Automatic Initialization):**
  - User opens the newly created file "Johnson Residence Phase 1 - Budget Sheet" from their Google Drive.
  - The `onOpen` script in this copy runs and identifies it as a new, uninitialized copy (because `IS_MASTER_TEMPLATE` is `'true'` and `MASTER_TEMPLATE_ACTUAL_ID` doesn't match the current file's ID, and `PROJECT_INITIALIZED` is not `'true'`).
  - The script automatically performs the following actions *without further user prompts for these specific items* (extracts "Johnson Residence Phase 1" as the Project Name from the file name, reads the `DATA_SHEET_ID` that was inherited from the Master Template, sets `PROJECT_NAME` to "Johnson Residence Phase 1", sets `PROJECT_INITIALIZED` to `'true'`, sets `IS_MASTER_TEMPLATE` to `'false'`, deletes the `MASTER_TEMPLATE_ACTUAL_ID` property, and sets `DATA_SHEET_ID` to the inherited value).
  - An alert appears: "Project 'Johnson Residence Phase 1' has been initialized successfully. It will use Data Sheet ID: [the inherited DATA_SHEET_ID]. The standard project menus are now available."
  - The full set of project-specific menus (e.g., 'Open Dashboard', 'Edit Project Configuration') are now loaded by `onOpen`.
- **User Opening an Already Initialized Project (Normal Operation):**
  - User re-opens "Johnson Residence Phase 1 - Budget Sheet" at a later time.
  - The `onOpen` script identifies it as an initialized project (`PROJECT_INITIALIZED` is `'true'`).
  - The standard project menus are loaded. The UI, when opened, will display "Johnson Residence Phase 1" as the project name and use the configured `DATA_SHEET_ID` for its operations.
- **User Opening an Unconfigured Sheet (e.g., a blank sheet or a copy made outside the script):**
  - User opens a Google Sheet that is not the Master Template and has not been through the script's initialization process.
  - The `onOpen` script finds no `IS_MASTER_TEMPLATE='true'` property (or it's false) and no `PROJECT_INITIALIZED='true'` property.
  - A limited menu is displayed, primarily offering an option like `Project Manager > Setup > Initialize Sheet as Project Manually...`.

### 5.2. Core experience
- **Step 1: Launch Project Manager UI**
  - User clicks the designated menu item (e.g., "Open Dashboard" or potentially an item under a "Project Manager" main menu).
  - *Good Experience:* The UI (dialog with a sidebar) loads quickly, displaying the current Project Name. The sidebar allows navigation between Rooms, Master Lists (if editing masters is part of this UI), etc.
- **Step 2: Manage/Select Rooms for the Current Project**
  - The UI displays a list of rooms currently selected for the project (initially empty or loaded from `_TempSelectedRooms`).
  - Users can add rooms by selecting from the Master Room List (sourced from Master Data Sheet) or by directly adding a new room name (which also adds to the Master Room List if it doesn't exist).
  - Users can remove rooms from the current project.
  - *Good Experience:* Clear visual distinction between adding/selecting master rooms and rooms in the current project. Easy search or filter for the Master Room List if it's long. Immediate feedback as rooms are added/removed from the project.
- **Step 3: Assign/Manage Category Types for a Selected Project Room**
  - User selects a room in their current project list.
  - The UI displays category types currently assigned to this room (initially empty or loaded from `_TempRoomTypes` for that room).
  - Users can assign category types by selecting from the Master Category Types List.
  - Users can remove a category type from the current room.
  - *Good Experience:* Context is clear (selected room is highlighted). Master Category Types are easy to browse/select. The interface is presented in the main content area for better visibility and space.
- **Step 4: Add/Manage Items for a Selected Category within a Project Room**
  - User selects a category type within a selected project room.
  - The UI displays items currently added to this room/category (initially empty or loaded from `_TempItemSelections`).
  - Users can add items by selecting from the Master Items List (filtered by the chosen `Item-Type`/Category Type).
  - For each item added to the project, the user must specify the `QUANTITY`.
  - Users can edit the `QUANTITY` of an item or remove an item from this room/category.
  - *Good Experience:* Master Items list is easily searchable/filterable. Inputting quantity is straightforward. Clear indication of which items are already added.
- **Step 5: Repeat for other Rooms and Categories**
  - User navigates through their selected project rooms and categories, adding items and quantities as needed.
  - *Good Experience:* Consistent UI and interaction patterns across different rooms/categories. Easy navigation back and forth.
- **Step 6: Generate/Update Project Sheet**
  - User clicks a "Generate Sheet" or "Update Sheet" button in the UI.
  - The script populates/updates the designated output tab in the Active Project Sheet with columns: `ROOM`, `TYPE`, `ITEM`, `QUANTITY`, and blank columns for manual entry: `LOW`, `LOW TOTAL`, `HIGH`, `HIGH TOTAL`, `SPEC/FFE`.
  - *Good Experience:* Clear confirmation message upon successful sheet generation/update. Option to automatically open/navigate to the populated sheet.

### 5.3. Advanced features & edge cases
- None identified for V1.0 beyond standard error handling (e.g., graceful message if the Master Data Sheet is temporarily inaccessible or if `ScriptProperties` are missing).
- Consideration for very long master lists (Rooms, Categories, Items): UI should ideally provide search/filter capabilities for efficient selection.

### 5.4. UI/UX highlights
- The primary interface is a dialog with a sidebar, leveraging an existing design for familiarity and clear navigation between sections (e.g., Project Rooms, Master Data Management if included).
- Clean, uncluttered, and intuitive interface, minimizing clicks and cognitive load.
- Responsive behavior within the Google Sheets dialog/sidebar environment.
- Consistent visual styling and interaction patterns throughout the tool.
- Clear visual feedback for user actions (selections, additions, deletions, errors).

## 6. Narrative
Sarah, a Lead Interior Designer, wants to quickly create accurate and professional-looking item specification sheets for her client presentations because manually compiling this data is time-consuming and prone to errors. She uses the 'Norton Project Sheet' tool within Google Sheets. Because the tool allows her to select rooms, then categories, and then items from pre-defined master lists, and specify quantities all within an intuitive UI, she can rapidly build a complete project list. The ability to instantly generate a formatted Google Sheet saves her hours on each project, ensures consistency, and lets her focus more on design and less on data entry.

## 7. Success metrics
### 7.1. User-centric metrics
   - (Skipped for V1.0)
### 7.2. Business metrics
   - (Skipped for V1.0)
### 7.3. Technical metrics
   - (Skipped for V1.0)

## 8. Technical considerations
### 8.1. Integration points
   - Primarily Google Sheets (Active Project Sheet, Master Data Sheet) and Google Apps Script services (including `SpreadsheetApp`, `ScriptProperties`).
   - No other external system or advanced Google Workspace API integrations are planned for V1.0.

### 8.2. Data storage & privacy
   - All data (master lists, project-specific selections, generated output) is stored within user-owned Google Sheets.
   - Access control and privacy are managed through standard Google Workspace sharing settings for the respective Google Sheets.
   - The Master Data Sheet and temporary data sheets (`_Temp...` sheets in the Active Project Sheet) will be hidden and protected where possible to minimize accidental user modification.
   - No specific advanced data validation mechanisms are planned for V1.0 beyond the UI-driven management of master lists and project selections.

### 8.3. Scalability & performance
   - Anticipated master data scale: Approximately 30 rooms, 50 category types, and 150 items.
   - Typical project item counts are expected to be well within a range manageable by Google Apps Script performance capabilities.
   - Script design should favor efficient data handling (e.g., batch reads/writes to sheets, minimizing calls within loops) to maintain UI responsiveness and stay within Apps Script execution time limits.
   - UI components displaying master lists (especially items) should be designed to handle ~150 entries efficiently (e.g., through client-side filtering if all data is loaded, or optimized server-side filtering if lists are fetched dynamically).

### 8.4. Potential challenges
   - Accidental user modification/corruption of the Master Data Sheet or the hidden temporary sheets (`_Temp...`) in the Active Project Sheet if users unhide/unprotect them. Mitigation includes hiding and protecting these sheets.
   - Concurrent use of the App Script UI by multiple users on the same Active Project Sheet simultaneously could potentially lead to race conditions or conflicts with `ScriptProperties` or temporary sheet data. This scenario is considered unlikely for V1.0.
   - Adherence to Google Apps Script quotas and limitations (e.g., script execution time, UI complexity, custom function limits) must be considered during development.

### 8.5. Code Documentation
   - **Inline JSDoc Comments:** All server-side JavaScript functions (`.js` files) are documented using JSDoc-style comments. This includes descriptions of the function's purpose, parameters (`@param`), return values (`@returns`), and any thrown exceptions (`@throws`). This is the primary source for detailed function-level documentation.
   - **README.md:** The `README.md` file provides a high-level overview of the project and a summary of each major script file's purpose.
   - **PRD (This Document):** This Product Requirements Document outlines the features, user stories, and overall functionality of the application.

## 9. Milestones & sequencing
### 9.1. Refactoring & Enhancement Focus Areas
   - This project focuses on refactoring and enhancing an existing codebase rather than distinct development phases.
   - **Focus Area 1: Data Management Logic (Ongoing Priority)**
     - Goal: Ensure all functions interacting with the Master Data Sheet and temporary sheets (`_Temp...`) are robust, efficient, and consistently handle data according to the defined architecture (e.g., `getRooms`, `getTypes`, `getAvailableItems`, and functions that save/delete data).
     - Key activities: Review and update data retrieval functions, data saving/update functions for master lists, and management of temporary selection sheets.
   - **Focus Area 2: Room Management UI & Logic (Current High Priority)**
     - Goal: Enhance the UI for managing rooms, particularly for editing and deleting individual rooms when displayed in the main content area. Improve clarity for adding, editing, and deleting rooms from the current project and master list.
     - Key activities: Refactor UI components for room display and interaction. Ensure logic correctly updates `_TempSelectedRooms` and the Master Data Sheet.
   - **Focus Area 3: Category Type Management UI & Logic (Future Enhancement Area)**
     - Goal: Streamline UI for assigning category types to rooms and managing the master list of category types.
     - Key activities: Review and update UI for category selection and master list management.
   - **Focus Area 4: Item Management & Selection UI & Logic (Future Enhancement Area)**
     - Goal: Improve the user experience for selecting items from the master list and setting quantities for the current project.
     - Key activities: Enhance UI for item browsing/searching/filtering. Ensure quantity input is intuitive.
   - **Focus Area 5: Sheet Generation Logic (Review/Refine Priority)**
     - Goal: Verify that the sheet generation process accurately reflects the data in the temporary sheets and produces the output in the specified column format (`ROOM`, `TYPE`, `ITEM`, `QUANTITY` + headers for manual entry columns).
     - Key activities: Review and optimize the function that populates the final output tab.

## 8. Changelog
### Version 1.4.2 (YYYY-MM-DD)
- Added functionality to copy specific columns from the generated "FFE" sheet to a "Pricing" sheet. The copied columns are ROOM, TYPE, ITEM, QUANTITY, LOW TOTAL, HIGH TOTAL from FFE, mapped to Room, Item Type, Item Name, Quantity, Budget Low, Budget High in the Pricing sheet.

### Version 1.4.1 (YYYY-MM-DD)
- **Fix**: Corrected an issue in the `_processAndCopyItemsInternal` function where formulas for "Low Total" and "High Total" columns in the "SPEC" sheet were referencing incorrect cells due to an additional "ACTUAL PRICE" column. The formulas are now dynamically generated using A1 notation based on the "SPEC" sheet's specific column layout for "QUANTITY", "LOW", and "HIGH" to ensure accurate calculations.

---
[Previous content of PRD continues from here if any]

<!-- ## 10. User stories

### 10.1. Set Project Name
- **ID**: US-002
- **Description**: As a Lead Interior Designer, I want to set a specific Project Name for my current design project so that I can easily identify it and the generated output reflects this name.
- **Acceptance Criteria**:
    - The UI provides a mechanism to input or update the Project Name.
    - The Project Name is saved to `ScriptProperties` of the Active Project Sheet.
    - The UI displays the current Project Name (e.g., in the dialog header).

### 10.2. View Master Room List
- **ID**: US-003
- **Description**: As a Lead Interior Designer, I want to view the list of all available master rooms so that I can choose which ones to include in my current project or manage the master list.
- **Acceptance Criteria**:
    - The UI displays a list of all room names sourced from the Master Data Sheet.
    - The list is clearly presented and easy to read.
    - If the list is very long, pagination or a search/filter mechanism is provided for ease of navigation.

### 10.3. Add a New Room to Master List
- **ID**: US-004
- **Description**: As a Lead Interior Designer, I want to add a new room name to the Master Room List so that it can be used in current and future projects.
- **Acceptance Criteria**:
    - The UI provides an input field and a button/action to add a new room name.
    - Upon submission, the new room name is added to the Master Room List in the Master Data Sheet.
    - The UI display of the Master Room List refreshes to show the newly added room.
    - Feedback (success/error) is provided.
    - Duplicate room names in the master list are handled gracefully (e.g., prevented with a message).

### 10.4. Edit an Existing Room in Master List
- **ID**: US-005
- **Description**: As a Lead Interior Designer, I want to edit the name of an existing room in the Master Room List so that I can correct typos or update naming conventions.
- **Acceptance Criteria**:
    - The UI allows selection of a room from the Master Room List for editing.
    - Upon saving changes, the room name is updated in the Master Data Sheet.
    - The UI display of the Master Room List refreshes to show the updated room name.
    - Feedback (success/error) is provided.

### 10.5. Delete a Room from Master List
- **ID**: US-006
- **Description**: As a Lead Interior Designer, I want to delete a room from the Master Room List so that obsolete rooms are removed from availability.
- **Acceptance Criteria**:
    - The UI allows selection of a room from the Master Room List for deletion.
    - A confirmation prompt is displayed before deletion.
    - Upon confirmation, the room is removed from the Master Data Sheet.
    - If the room is currently selected in the `_TempSelectedRooms` sheet of the Active Project Sheet, it is also removed from there.
    - The UI display of the Master Room List refreshes.
    - Feedback (success/error) is provided.

### 10.6. Select Rooms for Current Project
- **ID**: US-007
- **Description**: As a Lead Interior Designer, I want to select rooms from the Master Room List to include in my current design project so that I can define the scope of my project.
- **Acceptance Criteria**:
    - The UI displays the Master Room List allowing for multi-selection or individual selection of rooms.
    - Selected rooms are added to a "Current Project Rooms" list visible in the UI.
    - The names of selected rooms are stored in the `_TempSelectedRooms` sheet in the Active Project Sheet.
    - User can de-select a room from the Master List to remove it from the Current Project list (and `_TempSelectedRooms`).

### 10.7. View Rooms Selected for Current Project
- **ID**: US-008
- **Description**: As a Lead Interior Designer, I want to clearly see the list of rooms currently selected for my active design project so that I can manage their categories and items.
- **Acceptance Criteria**:
    - The UI prominently displays a list of rooms that are part of the current project (sourced from `_TempSelectedRooms`).
    - This list is the basis for further actions like assigning categories and items.

### 10.8. Remove a Room from Current Project
- **ID**: US-009
- **Description**: As a Lead Interior Designer, I want to remove a room from my current design project (without deleting it from the master list) so that I can adjust the project scope.
- **Acceptance Criteria**:
    - The UI allows selecting a room from the "Current Project Rooms" list.
    - Upon action (e.g., click a 'remove' icon), the room is removed from the "Current Project Rooms" list in the UI.
    - The room is removed from the `_TempSelectedRooms` sheet.
    - The room remains in the Master Room List.
    - Feedback (success/error) is provided.

### 10.9. View Master Category Types List
- **ID**: US-009
- **Description**: As a Lead Interior Designer, I want to view the list of all available master category types so that I can choose which ones to assign to rooms or manage the master list.
- **Acceptance Criteria**:
    - The UI displays a list of all category type names sourced from the Master Data Sheet.
    - The list is clearly presented and easy to read.
    - If the list is very long, pagination or a search/filter mechanism is provided.

### 10.10. Add a New Category Type to Master List
- **ID**: US-010
- **Description**: As a Lead Interior Designer, I want to add a new category type to the Master Category Types List so that it can be used for categorizing items.
- **Acceptance Criteria**:
    - The UI provides an input field and a button/action to add a new category type name.
    - Upon submission, the new category type is added to the Master Category Types List in the Master Data Sheet.
    - The UI display of the Master Category Types List refreshes to show the newly added category type.
    - Feedback (success/error) is provided.
    - Duplicate category type names in the master list are handled gracefully (e.g., prevented with a message).

### 10.11. Edit an Existing Category Type in Master List
- **ID**: US-011
- **Description**: As a Lead Interior Designer, I want to edit the name of an existing category type in the Master Category Types List so that I can correct typos or update naming conventions.
- **Acceptance Criteria**:
    - The UI allows selection of a category type from the Master Category Types List for editing.
    - Upon saving changes, the category type name is updated in the Master Data Sheet.
    - The UI display of the Master Category Types List refreshes to show the updated name.
    - Feedback (success/error) is provided.

### 10.12. Delete a Category Type from Master List
- **ID**: US-012
- **Description**: As a Lead Interior Designer, I want to delete a category type from the Master Category Types List so that obsolete categories are removed.
- **Acceptance Criteria**:
    - The UI allows selection of a category type from the Master Category Types List for deletion.
    - A confirmation prompt is displayed before deletion.
    - **Deletion is prevented if the Category Type is currently assigned as an `Item-Type` to any items in the Master Items List. An informative message is displayed to the user in this case.**
    - If not in use by Master Items, upon confirmation, the category type is removed from the Master Data Sheet.
    - If the category type was selected in the `_TempRoomTypes` sheet of the Active Project Sheet, it is also removed from there.
    - The UI display of the Master Category Types List refreshes.
    - Feedback (success/error) is provided.

### 10.13. View Assignable Category Types for a Project Room
- **ID**: US-013
- **Description**: As a Lead Interior Designer, when I have a room selected in my current project, I want to see the list of Master Category Types so that I can assign them to this room.
- **Acceptance Criteria**:
    - When a room is selected in the "Current Project Rooms" list, the UI displays the Master Category Types List (potentially in a separate panel or section).
    - The list allows for selection of one or more category types to be assigned to the currently selected project room.

### 10.14. Assign Category Type(s) to a Project Room
- **ID**: US-014
- **Description**: As a Lead Interior Designer, I want to assign one or more Master Category Types to a specific room in my current project so that I can organize items within that room.
- **Acceptance Criteria**:
    - User can select one or more category types from the Master Category Types List to assign to the currently active project room.
    - Upon assignment, the selected category type(s) are associated with the specific project room in the `_TempRoomTypes` sheet.
    - The UI updates to show the category types now assigned to that room (e.g., a sub-list under the room, or a tag display).
    - Feedback (success/error) is provided.

### 10.15. View Assigned Category Types for a Project Room
- **ID**: US-015
- **Description**: As a Lead Interior Designer, I want to clearly see which category types have been assigned to each specific room in my current project so that I can manage items under them.
- **Acceptance Criteria**:
    - When a room is selected in the "Current Project Rooms" list, the UI clearly displays the category types that have already been assigned to it (sourced from `_TempRoomTypes`).
    - This list of assigned category types is distinct from the master list and forms the basis for item addition within that room.

### 10.16. Remove an Assigned Category Type from a Project Room
- **ID**: US-016
- **Description**: As a Lead Interior Designer, I want to remove an assigned category type from a specific room in my current project (without deleting it from the master list) so that I can refine the project's structure.
- **Acceptance Criteria**:
    - The UI allows selecting an assigned category type from a specific project room.
    - Upon action (e.g., click a 'remove' icon), the category type is disassociated from that room in the UI and in the `_TempRoomTypes` sheet.
    - Any items previously added under this category type in this specific room (in `_TempItemSelections`) are also removed.
    - The category type remains in the Master Category Types List.
    - Feedback (success/error) is provided.

### 10.17. View Master Items List
- **ID**: US-017
- **Description**: As a Lead Interior Designer, I want to view the list of all available master items, along with their associated Item Types, so that I can manage this central repository and select items for projects.
- **Acceptance Criteria**:
    - The UI displays a list of all master items sourced from the Master Data Sheet, showing both `Item-Name` and `Item-Type`.
    - The list is clearly presented and easy to read.
    - If the list is very long (e.g., >150 items), pagination or a search/filter mechanism (e.g., by Item-Name or Item-Type) is provided.

### 10.18. Add a New Item to Master List
- **ID**: US-018
- **Description**: As a Lead Interior Designer, I want to add a new item (defining its Name and assigning its Type) to the Master Items List so that it can be used in projects.
- **Acceptance Criteria**:
    - The UI provides input fields for `Item-Name` and a way to select an `Item-Type` (from the Master Category Types List).
    - Upon submission, the new item (with its type and name) is added to the Master Items List in the Master Data Sheet.
    - The UI display of the Master Items List refreshes to show the newly added item.
    - Feedback (success/error) is provided.
    - Duplicate `Item-Name` within the same `Item-Type` in the master list are handled gracefully (e.g., prevented with a message, or allowed if `Item-Name` uniqueness is global).

### 10.19. Edit an Existing Item in Master List
- **ID**: US-019
- **Description**: As a Lead Interior Designer, I want to edit the Name and/or Type of an existing item in the Master Items List so that I can correct details or re-categorize items.
- **Acceptance Criteria**:
    - The UI allows selection of an item from the Master Items List for editing.
    - User can modify the `Item-Name` and re-assign the `Item-Type`.
    - Upon saving changes, the item details are updated in the Master Data Sheet.
    - The UI display of the Master Items List refreshes.
    - Feedback (success/error) is provided.

### 10.20. Delete an Item from Master List
- **ID**: US-020
- **Description**: As a Lead Interior Designer, I want to delete an item from the Master Items List so that obsolete items are removed from availability.
- **Acceptance Criteria**:
    - The UI allows selection of an item from the Master Items List for deletion.
    - A confirmation prompt is displayed before deletion.
    - Upon confirmation, the item is removed from the Master Data Sheet.
    - If the item is currently selected in the `_TempItemSelections` sheet of the Active Project Sheet, it is also removed from there.
    - The UI display of the Master Items List refreshes.
    - Feedback (success/error) is provided.

### 10.21. View Filtered Master Items for Project Selection
- **ID**: US-021
- **Description**: As a Lead Interior Designer, when I have selected a room and a category type within that room for my current project, I want to see a list of Master Items filtered by that category type so that I can select relevant items.
- **Acceptance Criteria**:
    - When a room and a category type within that room are active in the UI, the UI displays a list of `Item-Name`s from the Master Items List whose `Item-Type` matches the selected category type.
    - The list is easy to browse, and if long, offers search/filter by `Item-Name`.

### 10.22. Add Master Item to Current Project Room/Category
- **ID**: US-022
- **Description**: As a Lead Interior Designer, I want to add an item from the filtered Master Items list to the currently selected room and category in my project, and then specify its quantity.
- **Acceptance Criteria**:
    - User can select an item from the filtered Master Items list.
    - Upon selection, the item (identified by its Master `Item-Name` and `Item-Type`) is associated with the current project room/category in the `_TempItemSelections` sheet.
    - The UI prompts the user to input the `QUANTITY` for this item instance.
    - The `QUANTITY` is saved with the item instance in `_TempItemSelections`.
    - The UI updates to show the item now added to the current room/category, along with its quantity.
    - Feedback (success/error) is provided.

### 10.23. Edit Quantity of an Item in Current Project
- **ID**: US-023
- **Description**: As a Lead Interior Designer, I want to edit the quantity of an item that has already been added to a room/category in my current project so that I can make adjustments.
- **Acceptance Criteria**:
    - The UI allows selection or direct editing of the `QUANTITY` for an item listed within a project room/category.
    - Upon change, the updated `QUANTITY` is saved in the `_TempItemSelections` sheet for that item instance.
    - The UI display of the item's quantity updates immediately.
    - Feedback (success/error) is provided.

### 10.24. View Items Added to Current Project Room/Category
- **ID**: US-024
- **Description**: As a Lead Interior Designer, I want to clearly see the list of items (and their quantities) that have been added to each specific room and category in my current project.
- **Acceptance Criteria**:
    - When a project room and category are selected, the UI displays a list of items added to it, showing `Item-Name` and current `QUANTITY` (sourced from `_TempItemSelections`).
    - This list is clearly presented and forms the basis for quantity adjustments or item removal.

### 10.25. Remove an Item from Current Project Room/Category
- **ID**: US-025
- **Description**: As a Lead Interior Designer, I want to remove an item from a specific room/category in my current project (without deleting it from the master list) so that I can change my design selections.
- **Acceptance Criteria**:
    - The UI allows selecting an item from the list of items added to a project room/category.
    - Upon action (e.g., click a 'remove' icon), the item is removed from association with that room/category in the UI and in the `_TempItemSelections` sheet.
    - The item remains in the Master Items List.
    - Feedback (success/error) is provided.

### 10.26. Generate Project Specification Sheet
- **ID**: US-026
- **Description**: As a Lead Interior Designer, after configuring my project rooms, categories, items, and quantities, I want to generate a formatted specification sheet in my Active Project Google Sheet so that I have a deliverable output for clients or internal use.
- **Acceptance Criteria**:
    - The UI provides a button/action (e.g., "Generate Sheet" or "Update Sheet").
    - Upon action, the script reads the current project configuration from the temporary sheets (`_TempSelectedRooms`, `_TempRoomTypes`, `_TempItemSelections`).
    - A designated output tab in the Active Project Sheet is cleared (if it exists) or created.
    - The output tab is populated with the project data, grouped by Room, then by Category (Type).
    - The script populates the columns: `ROOM`, `TYPE`, `ITEM`, `QUANTITY` based on the current project configuration.
    - The script also creates headers for the columns intended for manual user input: `LOW`, `LOW TOTAL`, `HIGH`, `HIGH TOTAL`, `SPEC/FFE`.
    - The process is reasonably fast for a typical project size.
    - Feedback (success/error, e.g., "Sheet generated successfully") is provided to the user.
    - (Optional Good-to-have): The user is informed of the name of the output tab, or the script navigates the user to the generated/updated tab.

 --> 