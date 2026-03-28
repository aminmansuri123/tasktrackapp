# Office Task Management System - HTML Version

A simple, standalone HTML-based task management application that requires **NO INSTALLATION**. Just open the HTML file in your web browser!

## Features

✅ **No Installation Required** - Just open `index.html` in any modern web browser  
✅ **Works Offline** - All data stored in browser's localStorage  
✅ **Complete Task Management** - Create, edit, assign, and track tasks  
✅ **One-time and Recurring Tasks** - Support for both task types  
✅ **Working Days Calculation** - Automatically excludes weekends and holidays  
✅ **User Management** - Admin can add/edit/disable users  
✅ **Dashboard** - View statistics and today's tasks  
✅ **Calendar View** - Visual calendar with all tasks  
✅ **Export Functionality** - Export tasks to CSV or JSON  
✅ **Settings Management** - Manage locations, segregation types, and holidays  

## How to Use

1. **Open the Application**
   - Simply double-click `index.html` or open it in any web browser (Chrome, Firefox, Edge, Safari)
   - No server, no installation, no setup required!

2. **First Time Setup**
   - The first user to register automatically becomes an **Admin**
   - Admin users can:
     - Manage other users (add/edit/disable)
     - Manage settings (locations, segregation types, holidays)
     - View all tasks (team and self)

3. **Regular Users**
   - Can create and manage their own tasks
   - Can view team tasks (if assigned)
   - Can view dashboard and calendar

## Features Details

### Task Management
- **Task Name** (required)
- **Description**
- **Assigned To** (dropdown of users)
- **Location** (dropdown: Mundra, JNPT, Combine - can add more)
- **Task Type**: One-time or Recurring
- **Due Date** (for one-time tasks)
- **Priority** (High/Medium/Low for one-time tasks)
- **Recurrence Settings** (for recurring tasks):
  - Calendar Day or Working Day
  - Recurrence interval
- **Expected Completion Date**
- **Segregation Type** (PSA Reports, Internal Reports - can add more)
- **Estimated Minutes**
- **Team/Self Task** toggle
- **Task Status**: Not Completed, Completed, Completed but Need Improvement
- **Comments**

### Recurring Tasks
- **Calendar Day**: Updates automatically based on calendar days
- **Working Day**: Excludes weekends and configured holidays
- When a recurring task is marked as completed, the next due date is automatically calculated

### Dashboard
- Today's tasks count
- Overdue tasks count
- Completed tasks (last 30 days)
- Pending tasks count
- List of today's tasks

### Calendar View
- Monthly calendar view
- Tasks displayed on their due dates
- Color-coded by priority
- Navigate between months

### User Management (Admin Only)
- Add new users
- Edit user details
- Enable/disable users
- Change passwords
- Delete users

### Settings (Admin Only)
- **Locations**: Add/remove locations
- **Segregation Types**: Add/remove task segregation types
- **Holidays**: Add/remove holidays (used for working day calculations)

### Export
- Export tasks to CSV
- Export all data to JSON

## Data Storage

All data is stored in the browser's **localStorage**. This means:
- Data persists between browser sessions
- Data is specific to the browser/computer
- No server or database required
- Data can be exported for backup

## Browser Compatibility

Works on all modern browsers:
- Chrome (recommended)
- Firefox
- Edge
- Safari
- Opera

## File Structure

```
todo-list-app-html/
├── index.html    # Main application file (open this!)
├── app.js        # JavaScript logic
└── README.md     # This file
```

## Tips

1. **Backup Your Data**: Regularly export your data (CSV/JSON) to backup
2. **Browser Storage**: If you clear browser data, you'll lose all tasks. Export before clearing!
3. **Multiple Users**: Each user should use their own browser profile or export/import data
4. **Holidays**: Configure holidays in Settings to ensure accurate working day calculations

## Limitations

- Data is stored locally in the browser (not shared across devices)
- No server-side backup (export regularly!)
- No real-time collaboration (single user per browser)
- No email notifications
- No file attachments

## Troubleshooting

**Data Lost?**
- Check if browser data was cleared
- Try exporting data regularly as backup

**Can't Login?**
- Make sure you've registered first
- Check if account is disabled (admin can enable)

**Tasks Not Showing?**
- Check filters in Tasks tab
- Verify task assignment
- Check if viewing as admin (can see all tasks)

## Support

This is a standalone application. All features are self-contained in the HTML and JavaScript files.

---

**Enjoy managing your office tasks!** 🎉







