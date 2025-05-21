# PRD: Norton Project Sheet

## 1. Product overview
### 1.1 Document title and version
- PRD: Norton Project Sheet
- Version: 1.0

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
     - On first run or via a settings menu, allow user to set/update the ID of the Master Data Sheet (containing master lists for rooms, category types, and items) in `ScriptProperties` of the Active Project Sheet.
     - Allow user to set/update a "Project Name" (e.g., "Johnson Residence Living Room") for the current design project, stored in `ScriptProperties` of the Active Project Sheet.
     - The tool reads the Master Data Sheet ID (using `SpreadsheetApp.openById(ScriptProperties.getProperty('DATA_SHEET_ID'))`) and Project Name from `ScriptProperties` upon loading.
   - **Master Data Management (via UI interacting with Master Data Sheet)** (Priority: High)
     - **Master Rooms List (from Master Data Sheet):**
       - Allow user to add new room names to the master list in the Master Data Sheet (e.g., via `getRooms`, `getRoomNamesFromSheet`).
       - Allow user to edit existing room names in the master list.
       - When a room is deleted via the UI, it is removed from the Master Data Sheet and from the `_TempSelectedRooms` sheet in the Active Project Sheet.
     - **Master Category Types List (from Master Data Sheet):**
       - Allow user to add new category types (e.g., "Furniture," "Lighting") to the master list in the Master Data Sheet (e.g., via `getTypes`).
       - Allow user to edit names of existing category types in the master list.
       - When a category type is deleted via the UI, it is removed from the Master Data Sheet and from the `_TempRoomTypes` sheet in the Active Project Sheet. (Awaiting user decision on strategy if category type is used by Master Items).
     - **Master Items List (from Master Data Sheet):**
       - The Master Data Sheet contains an "Items" list/tab with columns: `Item-Type` (referencing a Master Category Type) and `Item-Name`.
       - Allow user to add new items (with their `Item-Type` and `Item-Name`) to the master list in the Master Data Sheet (e.g., via `getAvailableItems`).
       - Allow user to edit the `Item-Type` and `Item-Name` of existing items in the master list.
       - When an item is deleted via the UI, it is removed from the Master Data Sheet and from the `_TempItemSelections` sheet in the Active Project Sheet.
   - **Current Project Building (UI populating temporary sheets in Active Project Sheet)** (Priority: High)
     - **Room Selection for Current Project:**
       - Allow user to select rooms (from the Master Rooms List) to include in the current design project. Selections are stored temporarily in the `_TempSelectedRooms` sheet in the Active Project Sheet.
       - Display the list of selected rooms for the current project.
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
   - **User Interface (UI) Management** (Priority: High)
     - Provide a clear, intuitive, and responsive dialog UI for all management and selection functions.
     - UI should clearly distinguish between managing master data and building the current project.
     - Ensure UI elements are consistently styled and easy to understand.
     - Provide feedback to the user for actions (e.g., "Master Room added," "Item selected for project," "Error saving configuration").
## 5. User experience
### 5.1. Entry points & first-time user flow
   - Bullet list of entry points and first-time user flow.
### 5.2. Core experience
   - Step by step bullet list of the core experience in the following format:
   - **{step_1}**: {explanation_of_step_1}
     - {how_to_make_it_a_good_first_experience}
### 5.3. Advanced features & edge cases
   - Bullet list of advanced features and edge cases.
### 5.4. UI/UX highlights
   - Bullet list of UI/UX highlights.
## 6. Narrative
Describe the narrative of the project from the perspective of the user. For example: "{name} is a {role} who wants to {goal} because {reason}. {He/She} finds {project} and {reason_it_works_for_them}." Explain the users journey and the benefit they get from the end result. Limit the narrative to 1 paragraph only.
## 7. Success metrics
### 7.1. User-centric metrics
   - Bullet list of user-centric metrics.
### 7.2. Business metrics
   - Bullet list of business metrics.
### 7.3. Technical metrics
   - Bullet list of technical metrics.
## 8. Technical considerations
### 8.1. Integration points
   - Bullet list of integration points.
### 8.2. Data storage & privacy
   - Bullet list of data storage & privacy considerations.
### 8.3. Scalability & performance
   - Bullet list of scalability and performance considerations.
### 8.4. Potential challenges
   - Bullet list of potential challenges.
## 9. Milestones & sequencing
### 9.1. Project estimate
   - Bullet list of project estimate. i.e. "Medium: 2-4 weeks", eg:
   - {Small|Medium|Large}: {time_estimate}
### 9.2. Team size & composition
   - Bullet list of team size and composition. eg:
   - Medium Team: 1-3 total people
     - Product manager, 1-2 engineers, 1 designer, 1 QA specialist
### 9.3. Suggested phases
   - Bullet list of suggested phases in the following format:
   - **{Phase 1}**: {description_of_phase_1} ({time_estimate})
     - {key_deliverables}
   - **{Phase 2}**: {description_of_phase_2} ({time_estimate})
     - {key_deliverables}
## 10. User stories
Create a h3 and bullet list for each of the user stories in the following example format:
### 10.{x}. {user_story_title}
   - **ID**: {user_story_id}
   - **Description**: {user_story_description}
   - **Acceptance criteria**: {user_story_acceptance_criteria} 