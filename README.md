# outlook-task-creator

This repository contains a VBA script for Microsoft Outlook to automatically create a task from a selected email and assign a specific category.

## Environment Details

This script was developed and tested under the following environment:

- **Operating System**: Windows 10 Enterprise
- **Application**: Microsoft Office LTSC Professional Plus 2021
- **Specific Application Version**: Outlook Desktop Application


## How to Use

1. Open Microsoft Outlook.
2. Press `Alt + F11` to open the Visual Basic for Applications editor.
3. In the left pane, expand `Microsoft Outlook Objects`, and double-click on `ThisOutlookSession`.
4. In the right pane, paste the VBA script provided in this repo.
5. Close the VBA editor.
6. Add the macro to the ribbon or toolbar for easy access:
   - Right-click on the toolbar or ribbon in Outlook and choose "Customize Ribbon" or "Customize Quick Access Toolbar".
   - Choose "Macros" from the "Choose commands from" dropdown on the left.
   - Select the macro `ThisOutlookSession.CreateTaskFromEmail` from the list and add it to the desired location.
   - Click OK.

Now, when you select an email and run the script (either by clicking the macro from the toolbar/ribbon), it will create a task with the specified properties. You can then add notes or make other changes to the task as needed.

