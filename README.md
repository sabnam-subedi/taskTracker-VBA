# taskTracker-VBA

A smart task tracker system built in Excel using VBA. Designed to help teams manage departmental tasks efficiently, track progress visually, and dynamically assign responsibilities based on department and employee.

## Features

- Central task entry form with department-based routing
- Automatic syncing of tasks to department sheets
- Duplicate detection using Task ID
- Supports task metadata: due date, priority, status, remaining days
- Status-based color coding (To Do, In Progress, Done)
- Dynamic employee assignment using drop-downs filtered by department
- Modular VBA code exported for version control
- Scalable architecture with dedicated department and employee sheets

## Project Structure
```
taskTracker-VBA/
├── Excel/         # Contains the working taskTracker xlsm file
├── VBA/           # Contains exported .bas/.frm/.cls modules
├── Docs/          # Contains documentation
├── README.md      
└── .gitignore     
```


## Getting Started

1. Download the `taskTracker.xlsm` file from the `Excel/` directory
2. Open it in Excel and enable macros
3. Use the Enter the Task button to input a task or manually input in the main sheet
4. Click Sync Task to distribute manually entered tasks to department sheets

## Developer Guide

To view or contribute to the VBA code:

1. Open `taskTracker.xlsm` in Excel
2. Press `Alt + F11` to open the VBA Editor
3. Use File → Import to bring in `.bas`, `.frm`, or `.cls` files from the `/VBA/` folder

## Notes

- Each department has its own dedicated sheet. Tasks submitted via the form or synced manually are routed automatically.
- The "Assigned To" column in department sheets uses a drop-down list populated from the `Employee_List` sheet, dynamically filtered by department.
- Employees and departments can be added at any point, making the system easily scalable.

## Full Documentation

For a detailed explanation of features, screenshots, form fields, and logic:  
[View Full Docs (.docx)](https://github.com/sabnam-subedi/taskTracker-VBA/blob/main/Docs/taskTracker.docx)

## Future Enhancements

- Email notifications when an employee is assigned a task
- Dashboard integration (Power BI or Excel Pivot Visuals)
- Task dependencies and Gantt-style visual scheduling
- User-role-based access control
- Archived task handling

---

Built with VBA to organize tasks efficiently across departments.


