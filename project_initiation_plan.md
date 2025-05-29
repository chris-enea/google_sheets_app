# Project Initiation & Naming Convention - Development Plan

## 1. Overview
This plan outlines the tasks required to implement a refined project initiation process using a master template system and specific file naming conventions for Google Sheets managed by an Apps Script.

**Key Goals:**
-   Allow users to create new project sheets from a designated Master Template.
-   Ensure new project files are named `[Project Name] - Budget Sheet`.
-   Store the base "Project Name" in a script property for UI and logical use.
-   Streamline the setup process for new project copies by pre-configuring the `DATA_SHEET_ID`.

## 2. Initial Master Template Setup (Manual User/Admin Task - One Time)
-   [ ] **Action**: Manually access the Script Properties of the designated Master Template Google Sheet.
    -   Details: Via File > Project Properties > Script Properties in the Apps Script editor.
-   [ ] **Task**: Add/Set the script property `IS_MASTER_TEMPLATE` to the string value `'true'`.
-   [ ] **Task**: Add/Set the script property `MASTER_TEMPLATE_ACTUAL_ID` to the actual file ID of the Master Template sheet.
    -   Details: The file ID can be obtained from the URL of the Master Template sheet.
-   [ ] **Task**: Add/Set the script property `DATA_SHEET_ID` to your constant Master Data Sheet ID on the Master Template itself.
    -   Purpose: This ID will be inherited by and automatically applied to new projects created from this template.
-   [ ] **Task**: Ensure no project-specific properties (`PROJECT_NAME`, `PROJECT_INITIALIZED='true'`) exist on the Master Template, or that they are cleared. The `DATA_SHEET_ID` property, however, *should* exist on the master as it holds the default value.

## 3. Backend Development (Apps Script - `Code.js`)

### 3.1. `onOpen(e)` Function Enhancements
-   [x] **Task**: Refactor `onOpen(e)` to determine operational mode (`MASTER`, `NEW_PROJECT_COPY`, `INITIALIZED_PROJECT`, `UNCONFIGURED`) based on `IS_MASTER_TEMPLATE`, `MASTER_TEMPLATE_ACTUAL_ID`, `PROJECT_INITIALIZED` properties and the current file's ID.
    -   [x] Sub-Task: Implement logic for `MASTER` mode:
        -   [x] Verify `MASTER_TEMPLATE_ACTUAL_ID` matches `currentFileId`.
        -   [x] Ensure `PROJECT_INITIALIZED` is `false` (or absent) and project-specific properties like `PROJECT_NAME` are clear on the master template. (Verify `DATA_SHEET_ID` is present and correctly set on the master).
        -   [x] Prompt user: "Create NEW project" or "Edit Master Template".
        -   [x] If "Create NEW project":
            -   [x] Prompt for "Project Name".
            -   [x] Construct full file name: `[Project Name] - Budget Sheet`.
            -   [x] Copy the master sheet using `ss.copy(fullFileName)`. (All script properties, including the master's `DATA_SHEET_ID`, are copied).
            -   [x] Alert user with success message and link to the new file.
            -   [x] Return to prevent further menu loading in the template.
    -   [x] Sub-Task: Implement logic for `NEW_PROJECT_COPY` mode:
        -   [x] Condition: `IS_MASTER_TEMPLATE` is true (inherited), `MASTER_TEMPLATE_ACTUAL_ID` is present but *not* equal to `currentFileId`, and `PROJECT_INITIALIZED` is not true.
        -   [x] Extract "Project Name" from the current file name.
        -   [x] Alert user about new project initialization.
        -   [x] **Removed**: No longer prompt for "Master Data Sheet ID".
        -   [x] Read the inherited `DATA_SHEET_ID` from script properties.
        -   [x] If the inherited `DATA_SHEET_ID` is found and valid:
            -   [x] Set `PROJECT_NAME` script property to the extracted "Project Name".
            -   [x] (The `DATA_SHEET_ID` is already correctly set via inheritance from the master).
            -   [x] Set `PROJECT_INITIALIZED = 'true'`.
            -   [x] Set `IS_MASTER_TEMPLATE = 'false'`.
            -   [x] Delete `MASTER_TEMPLATE_ACTUAL_ID` from the project copy. (The `DATA_SHEET_ID` inherited from the master is kept).
            -   [x] Alert user of successful setup, confirming the Data Sheet ID used (which is the inherited one).
        -   [x] If the inherited `DATA_SHEET_ID` is missing or invalid (e.g., was not set on master):
            -   [x] Alert user of setup failure due to missing `DATA_SHEET_ID` on master template.
            -   [x] Revert mode to `UNCONFIGURED`.
    -   [x] Sub-Task: Implement logic for `INITIALIZED_PROJECT` mode:
        -   [x] Condition: `PROJECT_INITIALIZED` is true.
        -   [x] Perform self-correction: Ensure `IS_MASTER_TEMPLATE` is `false` and `MASTER_TEMPLATE_ACTUAL_ID` is deleted if found.
        -   [x] Log a warning and alert user if `DATA_SHEET_ID` is missing or empty.
    -   [x] Sub-Task: Implement logic for `UNCONFIGURED` mode.
        -   [x] Log if `IS_MASTER_TEMPLATE` is true but `MASTER_TEMPLATE_ACTUAL_ID` is missing.
-   [x] **Task**: Ensure `loadStandardMenus` is called correctly.

### 3.2. `loadStandardMenus(...)` Function Adjustments
    -   [x] (No changes directly related to `DATA_SHEET_ID` pre-fill, but ensure alerts for missing `DATA_SHEET_ID` in initialized projects are still relevant).

### 3.3. `saveProjectNameToProperties(projectName)` Function
-   [ ] (No changes needed from this specific requirement).

### 3.4. `initializeAsProjectManually(...)` Function
-   [x] **Task**: When prompting for "Project Name":
    -   [x] Suggest a default extracted from the current file name.
-   [x] **Task**: **Removed**: No longer prompt for "Master Data Sheet ID" initially.
-   [x] **Task**: Attempt to read `DATA_SHEET_ID` from script properties (this would be present if the sheet was copied from a master but its auto-initialization failed, or if it was manually set before).
    -   [x] If `DATA_SHEET_ID` is found and valid, confirm with the user if they want to use this existing ID.
    -   [x] If not found, or user declines, alert the user that `DATA_SHEET_ID` needs to be set. Provide a field to input it manually.
-   [x] **Task**: After getting valid "Project Name" and `DATA_SHEET_ID` (either confirmed inherited, or manually entered):
    -   [x] Set `PROJECT_NAME`, `DATA_SHEET_ID` (if manually entered or re-confirmed), `PROJECT_INITIALIZED='true'`, `IS_MASTER_TEMPLATE='false'`.
    -   [x] Delete `MASTER_TEMPLATE_ACTUAL_ID`.
    -   [x] Rename file and alert.

### 3.5. `editProjectProperties(...)` Function
-   [x] **Task**: Continue to allow editing of `DATA_SHEET_ID` in this function.
    -   This provides a way to override the default/inherited one or fix it if the `DATA_SHEET_ID` on the master was incorrect at the time of project creation, or if it needs to change post-initialization.

### 3.6. Removal of Obsolete Function
-   [ ] (No changes needed from this specific requirement).

## 4. Frontend / UI Considerations
-   [x] **Task**: Remove UI prompt for "Master Data Sheet ID" during the automatic new project copy initialization.
-   [x] **Task**: Update alert messages to confirm which `DATA_SHEET_ID` was used if set automatically.
-   [x] **Task**: Modify `initializeAsProjectManually` to reflect the new DATA_SHEET_ID handling (prefer automatic, fallback to prompt).

## 5. Testing Checklist
-   [ ] **Master Template Setup & Recognition**:
    -   [ ] Manually set `IS_MASTER_TEMPLATE`, `MASTER_TEMPLATE_ACTUAL_ID`, AND `DATA_SHEET_ID` on the master.
    -   [ ] Open Master: Verify it prompts "Create NEW project" or "Edit Master".
-   [ ] **New Project Creation**:
    -   [ ] From Master, choose "Create NEW project". Enter "Project Name".
    -   [ ] Verify new file `[Project Name] - Budget Sheet` is created.
-   [ ] **New Project Copy Initialization**:
    -   [ ] Open the newly created file.
    -   [ ] **Verify it does NOT prompt for "Master Data Sheet ID"**.
    -   [ ] Verify it correctly identifies the "Project Name".
    -   [ ] Verify script properties are set correctly in the new copy:
        -   `PROJECT_NAME` = (extracted name)
        -   `DATA_SHEET_ID` = (value inherited from master's `DATA_SHEET_ID`)
        -   `PROJECT_INITIALIZED` = 'true'
        -   `IS_MASTER_TEMPLATE` = 'false'
        -   `MASTER_TEMPLATE_ACTUAL_ID` is deleted/absent. (The `DATA_SHEET_ID` is present).
    -   [ ] Verify correct project-specific menus load.
    -   [ ] Test scenario: `DATA_SHEET_ID` is missing from Master Template – verify failure message on new copy initialization.
-   [ ] **Opening Initialized Project**:
    -   [ ] Verify `DATA_SHEET_ID` is correct (matches the one from master's `DATA_SHEET_ID`).
-   [ ] **`initializeAsProjectManually(...)`**:
    -   [ ] Test on an unconfigured sheet (that might have an inherited `DATA_SHEET_ID` if copied from master but init failed). Verify it offers to use it or prompts for a new one.
    -   [ ] Test on a truly blank unconfigured sheet (no inherited `DATA_SHEET_ID`). Verify it prompts for Data Sheet ID.
-   [ ] **`editProjectProperties(...)`**:
    -   [ ] Verify `DATA_SHEET_ID` can still be viewed and changed.
-   [ ] **Error Handling**:
    -   [ ] Test empty input for "Project Name".

## 6. Documentation Updates
-   [x] **Task**: Update `prd.md` to reflect that `DATA_SHEET_ID` is set on the master and inherited by new projects, removing the user prompt during automatic initialization.
-   [ ] **Task**: Update `README.md` and JSDoc comments.

## 7. Code Cleanup
-   [ ] (As before).

---
This plan should provide a good roadmap for implementing the changes. Let me know if you have any questions. 