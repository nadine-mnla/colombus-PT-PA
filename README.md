# COLOMBUS Hotel Management System

A legacy VB6 + SQL Server based hotel room management system originally developed in 2010. This project manages guest data, room availability, and displays occupancy rates.

## Modules & Forms

- `frmMain.frm` ? Room Chart UI
- `frmGuest.frm` ? Guest Entry UI
- `modDatabase.bas` ? Database connection logic
- `modRoomManagement.cls` ? Room status and occupancy logic

## How to Run

1. Open `Colombus.vbp` in Visual Basic 6.0
2. Run the project (`F5`)
3. Make sure SQL Server has a DB `ColombusDB` with the schema in `ColombusDB.sql`