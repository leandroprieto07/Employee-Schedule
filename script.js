// --- Simulation Data (NOT SECURE FOR PRODUCTION!) ---
// Defines users and their roles for the login system.
// IMPORTANT: 'users' is now a 'let' variable and is persisted in localStorage.
// The initial users are only used if localStorage is empty.
let users = JSON.parse(localStorage.getItem('users')) || {
    'admin': { password: 'adminpassword', role: 'admin' },
    'supervisor1': { password: 'sup1password', role: 'supervisor', displayName: 'Supervisor Alpha' } // Added displayName for default supervisor
};
// FOR DEBUGGING: Log the loaded users to the console
console.log('Loaded Users:', users);

// Load data from localStorage or initialize as empty
let currentUser = JSON.parse(localStorage.getItem('currentUser')) || null;
let employees = JSON.parse(localStorage.getItem('employees')) || [];
let calendarData = JSON.parse(localStorage.getItem('calendarData')) || {};

// --- DOM Elements ---
const loginContainer = document.getElementById('login-container');
const appContainer = document.getElementById('app-container');
const loginForm = document.getElementById('login-form');
const loginMessage = document.getElementById('login-message');
const userInfoSpan = document.getElementById('user-info');
const logoutButton = document.getElementById('logout-button');

const shiftsTable = document.getElementById('shifts-table');
const shiftsTableBody = document.getElementById('shifts-table-body');
const dateHeaderRow = document.getElementById('date-header');
const currentWeekRangeSpan = document.getElementById('current-week-range');
const prevWeekButton = document.getElementById('prev-week');
const nextWeekButton = document.getElementById('next-week');
const exportExcelButton = document.getElementById('export-excel-button');

const adminMainSection = document.getElementById('admin-main-section'); // New: Container for all admin sections
const employeeManagementSection = document.getElementById('employee-management-section');
const addEmployeeForm = document.getElementById('add-employee-form');
const employeeListBody = document.getElementById('employee-list-body');

// NEW: User Management DOM Elements
const userManagementSection = document.getElementById('user-management-section');
const addUserForm = document.getElementById('add-user-form');
const newUsernameInput = document.getElementById('new-username');
const newPasswordInput = document.getElementById('new-password');
const newSupervisorNameInput = document.getElementById('new-supervisor-name'); // This will store the supervisor's display name
const newUserRoleSelect = document.getElementById('new-user-role');
const userCreationMessage = document.getElementById('user-creation-message');
const userListBody = document.getElementById('user-list-body');


const statusModal = document.getElementById('status-modal');
const modalEmployeeName = document.getElementById('modal-employee-name');
const modalDate = document.getElementById('modal-date');
const statusSelect = document.getElementById('status-select');
const saveStatusButton = document.getElementById('save-status-button');
const closeButton = statusModal.querySelector('.close-button');

// --- Calendar State Variables ---
let startDate = new Date(); // Initialized with current date
const DISPLAY_DAYS = 14; // Number of days to display

// Variables for the editing modal
let currentEditingEmployeeId = null;
let currentEditingDate = null;
let currentEditingCell = null;


// --- Utility Functions ---

/**
 * Formats a Date object into a `YYYY-MM-DD` string.
 * @param {Date} date - The Date object to format.
 * @returns {string} The formatted date (e.g., "2023-01-15").
 */
function formatDateYYYYMMDD(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0'); // Months are 0-indexed
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

/**
 * Formats a `YYYY-MM-DD` date string to MM/DD/YYYY.
 * @param {string} dateStringYYYYMMDD - The date in `YYYY-MM-DD` format.
 * @returns {string} The formatted date (e.g., "01/15/2023").
 */
function formatDateMMDDYYYY(dateStringYYYYMMDD) {
    const [year, month, day] = dateStringYYYYMMDD.split('-');
    return `${month}/${day}/${year}`;
}

/**
 * Checks if a given date is a weekend (Saturday or Sunday).
 * @param {Date} date - The Date object to check.
 * @returns {boolean} True if it's a weekend, false otherwise.
 */
function isWeekend(date) {
    const dayOfWeek = date.getDay(); // 0 = Sunday, 6 = Saturday
    return dayOfWeek === 0 || dayOfWeek === 6;
}

// --- Login/Logout Functions ---

loginForm.addEventListener('submit', (e) => {
    e.preventDefault(); // Prevents page reload on form submission.
    const username = document.getElementById('username').value;
    const password = document.getElementById('password').value;

    if (users[username] && users[username].password === password) {
        currentUser = { username: username, role: users[username].role };
        // If the logged-in user is a supervisor, get their display name
        if (currentUser.role === 'supervisor' && users[username].displayName) {
            currentUser.displayName = users[username].displayName;
        } else if (currentUser.role === 'supervisor') {
             // Fallback: if a supervisor user doesn't have a displayName explicitly set, use their username
            currentUser.displayName = username;
        }
        localStorage.setItem('currentUser', JSON.stringify(currentUser));
        loginMessage.textContent = '';
        renderApp();
    } else {
        loginMessage.textContent = 'Incorrect username or password.';
    }
});

logoutButton.addEventListener('click', () => {
    currentUser = null; // Clears the current user.
    localStorage.removeItem('currentUser'); // Removes the user from localStorage.
    loginContainer.style.display = 'block'; // Shows the login container.
    appContainer.style.display = 'none'; // Hides the application container.
});

/**
 * Checks authentication status when the page loads.
 * If a user is found in localStorage, it logs them in automatically.
 */
function checkAuth() {
    const storedUser = localStorage.getItem('currentUser');
    if (storedUser) {
        currentUser = JSON.parse(storedUser);
        // Ensure displayName is set for currentUser if it's a supervisor during initial load
        if (currentUser.role === 'supervisor' && users[currentUser.username] && users[currentUser.username].displayName) {
            currentUser.displayName = users[currentUser.username].displayName;
        } else if (currentUser.role === 'supervisor') {
            currentUser.displayName = currentUser.username; // Default to username if no displayName set
        }
        renderApp();
    } else {
        loginContainer.style.display = 'block'; // Otherwise, show the login.
        appContainer.style.display = 'none';
    }
}

/**
 * Renders the main application interface after a successful login.
 * Shows/hides sections based on the user's role.
 */
function renderApp() {
    loginContainer.style.display = 'none'; // Hides the login form.
    appContainer.style.display = 'block'; // Shows the application interface.
    // Display supervisor's display name if available, otherwise username
    const displayUserName = currentUser.role === 'supervisor' && currentUser.displayName ? currentUser.displayName : currentUser.username;
    userInfoSpan.textContent = `Welcome, ${displayUserName} (${currentUser.role})`;

    if (currentUser.role === 'admin') {
        adminMainSection.style.display = 'block'; // Show the entire admin section
        employeeManagementSection.style.display = 'block';
        userManagementSection.style.display = 'block';
    } else {
        adminMainSection.style.display = 'none'; // Hide the entire admin section
        employeeManagementSection.style.display = 'none';
        userManagementSection.style.display = 'none';
    }

    renderShiftsCalendar();
    renderEmployeeList();
    renderUserList(); // Render user list on app load/login
}

// --- Horizontal Shift Calendar Functions ---

/**
 * Renders the horizontal shift calendar table.
 * Dynamically generates date headers and employee rows.
 */
function renderShiftsCalendar() {
    // Clear date headers and table body before rendering.
    dateHeaderRow.innerHTML = '';
    shiftsTableBody.innerHTML = '';

    // Adjust startDate to be the beginning of the week (Sunday)
    // Create a mutable copy of startDate for calculation
    const displayStartDate = new Date(startDate); 
    displayStartDate.setDate(displayStartDate.getDate() - displayStartDate.getDay()); // Set to Sunday of the current week

    const currentDay = new Date(displayStartDate); // Use this adjusted date for rendering
    const today = new Date();
    today.setHours(0,0,0,0);
    
    let firstDateDisplayed = null;
    let lastDateDisplayed = null;

    for (let i = 0; i < DISPLAY_DAYS; i++) {
        const dateString = formatDateYYYYMMDD(currentDay);
        const dayOfWeekName = currentDay.toLocaleString('en-US', { weekday: 'short' });
        const monthDay = currentDay.toLocaleString('en-US', { month: 'short', day: 'numeric' });

        const th = document.createElement('th');
        th.innerHTML = `${monthDay}<br>${dayOfWeekName}`;
        th.classList.add('date-col');

        if (isWeekend(currentDay)) {
            th.classList.add('weekend');
        }
        
        if (currentDay.setHours(0,0,0,0) === today.getTime()) {
            th.classList.add('today-column');
        }

        dateHeaderRow.appendChild(th);

        if (i === 0) firstDateDisplayed = new Date(currentDay);
        if (i === DISPLAY_DAYS - 1) lastDateDisplayed = new Date(currentDay);

        currentDay.setDate(currentDay.getDate() + 1);
    }

    if (firstDateDisplayed && lastDateDisplayed) {
        const start = firstDateDisplayed.toLocaleString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
        const end = lastDateDisplayed.toLocaleString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
        currentWeekRangeSpan.textContent = `${start} - ${end}`;
    }

    employees.forEach(employee => {
        const row = document.createElement('tr');
        row.dataset.employeeId = employee.id;

        row.innerHTML = `
            <td>${employee.area || ''}</td>
            <td>${employee.techNumber}</td>
            <td>${employee.firstName}</td>
            <td>${employee.lastName}</td>
            <td>${employee.supervisor || ''}</td> <!-- This will now display the supervisor's Display Name -->
            <td>${employee.status || 'Working'}</td>
        `;

        const employeeCalendar = calendarData[employee.id] || {};
        // Use the adjusted displayStartDate for employee rows too
        const tempDate = new Date(displayStartDate); 

        for (let i = 0; i < DISPLAY_DAYS; i++) {
            const dateString = formatDateYYYYMMDD(tempDate);
            const entry = employeeCalendar[dateString];
            
            let statusToDisplay;
            let cssClassToAdd;

            if (typeof entry === 'object' && entry.status === 'pending') {
                statusToDisplay = `PENDING (${entry.requestedStatus.toUpperCase()})`;
                cssClassToAdd = 'pending';
            } else if (typeof entry === 'string') {
                statusToDisplay = entry.toUpperCase();
                cssClassToAdd = entry;
            } else {
                statusToDisplay = 'WORKING';
                cssClassToAdd = 'working';
            }

            const cell = document.createElement('td');
            cell.textContent = statusToDisplay;
            cell.classList.add(cssClassToAdd);

            if (isWeekend(tempDate)) {
                cell.classList.add('weekend');
            }
            if (tempDate.setHours(0,0,0,0) === today.getTime()) {
                cell.classList.add('today-column');
            }

            cell.dataset.date = dateString;
            cell.dataset.employeeId = employee.id;

            // Handle cell click for editing based on user role and supervisor link (using displayName)
            if (currentUser) {
                if (currentUser.role === 'admin' || 
                   (currentUser.role === 'supervisor' && employee.supervisor === currentUser.displayName)) { // Use displayName for linkage check
                    cell.addEventListener('click', () => openStatusModal(employee, dateString, entry, cell));
                } else if (currentUser.role === 'supervisor' && employee.supervisor !== currentUser.displayName) {
                    cell.style.cursor = 'not-allowed';
                    cell.title = 'You can only manage shifts for your assigned employees.';
                }
            }

            row.appendChild(cell);
            tempDate.setDate(tempDate.getDate() + 1);
        }
        shiftsTableBody.appendChild(row);
    });
}

// --- Calendar Navigation ---

// Navigate to the previous `DISPLAY_DAYS` block, adjusting to the nearest Sunday start
prevWeekButton.addEventListener('click', () => {
    startDate.setDate(startDate.getDate() - DISPLAY_DAYS);
    renderShiftsCalendar();
});

// Navigate to the next `DISPLAY_DAYS` block, adjusting to the nearest Sunday start
nextWeekButton.addEventListener('click', () => {
    startDate.setDate(startDate.getDate() + DISPLAY_DAYS);
    renderShiftsCalendar();
});

// --- Day Status Modal ---

function openStatusModal(employee, dateString, currentEntry, cell) {
    // Basic permission check (using displayName)
    if (currentUser.role === 'supervisor' && employee.supervisor !== currentUser.displayName) {
        alert("You can only manage shifts for your assigned employees.");
        return;
    }

    currentEditingEmployeeId = employee.id;
    currentEditingDate = dateString;
    currentEditingCell = cell;

    modalEmployeeName.textContent = `${employee.firstName} ${employee.lastName} (Tech #: ${employee.techNumber})`;
    modalDate.textContent = formatDateMMDDYYYY(dateString);

    statusModal.querySelector('.modal-content h3').textContent = 'Change Day Status';
    const statusSelectContainer = document.getElementById('status-select-container');
    const adminActionsContainer = document.getElementById('admin-actions-container');
    
    adminActionsContainer.innerHTML = ''; 
    saveStatusButton.style.display = 'block';


    if (currentUser.role === 'admin' && typeof currentEntry === 'object' && currentEntry.status === 'pending') {
        statusModal.querySelector('.modal-content h3').textContent = 'Pending Approval Request';
        statusSelectContainer.style.display = 'none';
        saveStatusButton.style.display = 'none';

        adminActionsContainer.innerHTML = `
            <p>Requested Status: <strong>${currentEntry.requestedStatus.toUpperCase()}</strong></p>
            <p>Requested by: <strong>${currentEntry.requestedBy}</strong></p>
            <button id="approve-request-button" style="background-color: #28a745;">Approve</button>
            <button id="reject-request-button" style="background-color: #dc3545;">Reject</button>
        `;
        
        document.getElementById('approve-request-button').addEventListener('click', () => {
            if (!calendarData[currentEditingEmployeeId]) calendarData[currentEditingEmployeeId] = {};
            calendarData[currentEditingEmployeeId][currentEditingDate] = currentEntry.requestedStatus;
            localStorage.setItem('calendarData', JSON.stringify(calendarData));
            renderShiftsCalendar();
            statusModal.style.display = 'none';
            alert("Request approved.");
        });
        document.getElementById('reject-request-button').addEventListener('click', () => {
            if (!calendarData[currentEditingEmployeeId]) calendarData[currentEditingEmployeeId] = {};
            calendarData[currentEditingEmployeeId][currentEditingDate] = 'working'; 
            localStorage.setItem('calendarData', JSON.stringify(calendarData));
            renderShiftsCalendar();
            statusModal.style.display = 'none';
            alert("Request rejected.");
        });

    } else {
        statusSelectContainer.style.display = 'block';
        const initialStatusValue = typeof currentEntry === 'object' ? currentEntry.requestedStatus : (currentEntry || 'working');
        statusSelect.value = initialStatusValue;
    }

    statusModal.style.display = 'flex';
}

closeButton.addEventListener('click', () => {
    statusModal.style.display = 'none';
});

window.addEventListener('click', (event) => {
    if (event.target === statusModal) {
        statusModal.style.display = 'none';
    }
});

saveStatusButton.addEventListener('click', () => {
    const newStatus = statusSelect.value;
    const currentEntry = calendarData[currentEditingEmployeeId] ? calendarData[currentEditingEmployeeId][currentEditingDate] : null;
    const employeeToEdit = employees.find(emp => emp.id === currentEditingEmployeeId);

    // Re-check supervisor linkage for robustness (using displayName)
    if (currentUser.role === 'supervisor' && employeeToEdit.supervisor !== currentUser.displayName) {
        alert("You can only manage shifts for your assigned employees.");
        statusModal.style.display = 'none';
        return;
    }

    if (currentUser.role === 'supervisor') {
        if (!calendarData[currentEditingEmployeeId]) calendarData[currentEditingEmployeeId] = {};
        calendarData[currentEditingEmployeeId][currentEditingDate] = {
            status: 'pending',
            requestedStatus: newStatus,
            requestedBy: currentUser.displayName, // Store supervisor's display name
        };
        localStorage.setItem('calendarData', JSON.stringify(calendarData));
        statusModal.style.display = 'none';
        alert(`Request for ${newStatus.toUpperCase()} sent for approval.`);
        renderShiftsCalendar();

    } else if (currentUser.role === 'admin') {
        if (currentEntry && typeof currentEntry === 'object' && currentEntry.status === 'pending') {
            if (confirm(`ADMIN: Do you want to apply ${newStatus.toUpperCase()} directly (overriding pending)?`)) {
                if (!calendarData[currentEditingEmployeeId]) calendarData[currentEditingEmployeeId] = {};
                calendarData[currentEditingEmployeeId][currentEditingDate] = newStatus;
                localStorage.setItem('calendarData', JSON.stringify(calendarData));
                statusModal.style.display = 'none';
                alert("Status updated directly by Admin.");
                renderShiftsCalendar();
            } else {
                statusModal.style.display = 'none';
            }
        } else {
            if (!calendarData[currentEditingEmployeeId]) calendarData[currentEditingEmployeeId] = {};
            calendarData[currentEditingEmployeeId][currentEditingDate] = newStatus;
            localStorage.setItem('calendarData', JSON.stringify(calendarData));
            statusModal.style.display = 'none';
            alert("Status updated directly by Admin.");
            renderShiftsCalendar();
        }
    } else {
        alert("You do not have permissions to modify the status.");
        statusModal.style.display = 'none';
    }
});


// --- Employee Management Functions (Admin) ---

addEmployeeForm.addEventListener('submit', (e) => {
    e.preventDefault();
    if (currentUser.role !== 'admin') {
        alert("Only administrators can add employees.");
        return;
    }

    const area = document.getElementById('employee-area').value;
    const techNumber = document.getElementById('employee-tech-number').value;
    const firstName = document.getElementById('employee-first-name').value;
    const lastName = document.getElementById('employee-last-name').value;
    const supervisor = document.getElementById('employee-supervisor').value; // Supervisor's DISPLAY name

    // Validate that the entered supervisor display name exists as a supervisor
    const supervisorExists = Object.values(users).some(user => 
        user.role === 'supervisor' && user.displayName === supervisor
    );

    if (!supervisorExists && supervisor !== '') { // Allow empty supervisor if needed
        alert("The linked Supervisor Display Name does not exist or is not a supervisor user. Please ensure the display name matches an existing supervisor user.");
        return;
    }


    if (employees.some(emp => emp.techNumber === techNumber)) {
        alert("An employee with this Tech # already exists.");
        return;
    }

    const newEmployee = {
        id: Date.now().toString(),
        area,
        techNumber,
        firstName,
        lastName,
        supervisor, // Store the supervisor's display name
        status: 'Working'
    };
    employees.push(newEmployee);
    localStorage.setItem('employees', JSON.stringify(employees));
    addEmployeeForm.reset();
    renderEmployeeList();
    renderShiftsCalendar();
});

function renderEmployeeList() {
    employeeListBody.innerHTML = '';
    employees.forEach(emp => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${emp.area || ''}</td>
            <td>${emp.techNumber}</td>
            <td>${emp.firstName}</td>
            <td>${emp.lastName}</td>
            <td>${emp.supervisor || ''}</td> <!-- Display supervisor's display name -->
            <td>
                ${currentUser.role === 'admin' ? `<button onclick="editEmployee('${emp.id}')">Edit</button>
                                                   <button onclick="deleteEmployee('${emp.id}')">Delete</button>` : ''}
            </td>
        `;
        employeeListBody.appendChild(row);
    });
}

function editEmployee(id) {
    if (currentUser.role !== 'admin') {
        alert("Only administrators can edit employees.");
        return;
    }
    const employee = employees.find(emp => emp.id === id);
    if (employee) {
        const newArea = prompt("New Area:", employee.area);
        const newTechNumber = prompt("New Tech #:", employee.techNumber);
        const newFirstName = prompt("New First Name:", employee.firstName);
        const newLastName = prompt("New Last Name:", employee.lastName);
        const newSupervisor = prompt("New Supervisor (Display Name):", employee.supervisor); // Prompt for supervisor DISPLAY name
        
        // Validate that the entered supervisor display name exists as a supervisor
        const supervisorExists = Object.values(users).some(user => 
            user.role === 'supervisor' && user.displayName === newSupervisor
        );

        if (!supervisorExists && newSupervisor !== null && newSupervisor.trim() !== '') {
            alert("The linked Supervisor Display Name does not exist or is not a supervisor user. Please enter an existing supervisor's display name.");
            return; // Stop editing if invalid supervisor
        }

        if (newArea !== null) employee.area = newArea;
        if (newTechNumber !== null) employee.techNumber = newTechNumber;
        if (newFirstName !== null) employee.firstName = newFirstName;
        if (newLastName !== null) employee.lastName = newLastName;
        // Only update if not null and validated, or if it's explicitly being cleared
        if (newSupervisor !== null) employee.supervisor = newSupervisor.trim(); // Trim to ensure empty string if user clears it

        localStorage.setItem('employees', JSON.stringify(employees));
        renderEmployeeList();
        renderShiftsCalendar();
    }
}

function deleteEmployee(id) {
    if (currentUser.role !== 'admin') {
        alert("Only administrators can delete employees.");
        return;
    }
    if (confirm("Are you sure you want to delete this employee?")) {
        employees = employees.filter(emp => emp.id !== id);
        delete calendarData[id];
        localStorage.setItem('employees', JSON.stringify(employees));
        localStorage.setItem('calendarData', JSON.stringify(calendarData));
        renderEmployeeList();
        renderShiftsCalendar();
    }
}

// --- NEW: User Management Functions (Admin) ---

// Handle add user form submission
addUserForm.addEventListener('submit', (e) => {
    e.preventDefault();
    if (currentUser.role !== 'admin') {
        alert("Only administrators can manage users.");
        return;
    }

    const newUsername = newUsernameInput.value.trim();
    const newPassword = newPasswordInput.value.trim();
    const newUserRole = newUserRoleSelect.value;
    let newSupervisorDisplayName = newSupervisorNameInput.value.trim(); // Get potential display name

    userCreationMessage.textContent = ''; // Clear previous messages

    if (!newUsername || !newPassword) {
        userCreationMessage.textContent = 'Username and password cannot be empty.';
        return;
    }

    if (users[newUsername]) {
        userCreationMessage.textContent = `User '${newUsername}' already exists.`;
        return;
    }

    // Add the new user
    users[newUsername] = {
        password: newPassword,
        role: newUserRole
    };

    // If it's a supervisor role, store the display name. If empty, default to username.
    if (newUserRole === 'supervisor') {
        users[newUsername].displayName = newSupervisorDisplayName || newUsername;
    } else {
        // Ensure displayName is not set for admin users, or explicitly undefined
        delete users[newUsername].displayName;
    }
    
    localStorage.setItem('users', JSON.stringify(users)); // Persist users
    addUserForm.reset(); // Clear form
    newSupervisorNameInput.value = ''; // Ensure this is also cleared
    userCreationMessage.textContent = `User '${newUsername}' (${newUserRole}) created successfully.`;
    userCreationMessage.style.color = 'green';
    
    renderUserList(); // Update user list in UI
});

// Render the list of existing users
function renderUserList() {
    userListBody.innerHTML = '';
    for (const username in users) {
        const user = users[username];
        const row = document.createElement('tr');
        // Display supervisor's displayName if available
        const userDisplayInfo = user.role === 'supervisor' && user.displayName ? `(Display: ${user.displayName})` : '';
        row.innerHTML = `
            <td>${username}</td>
            <td>${user.role} ${userDisplayInfo}</td>
            <td>
                ${currentUser.role === 'admin' && username !== currentUser.username ? `
                    <button onclick="deleteUser('${username}')" style="background-color: #dc3545;">Delete</button>
                ` : ''}
            </td>
        `;
        userListBody.appendChild(row);
    }
}

// Delete a user (Admin only, cannot delete self)
function deleteUser(usernameToDelete) {
    if (currentUser.role !== 'admin') {
        alert("Only administrators can delete users.");
        return;
    }
    if (usernameToDelete === currentUser.username) {
        alert("You cannot delete your own user account.");
        return;
    }
    if (confirm(`Are you sure you want to delete user '${usernameToDelete}'? This cannot be undone.`)) {
        // Before deleting a user, check if any employees are linked to their displayName
        const userToDeleteDisplayName = users[usernameToDelete]?.displayName || usernameToDelete;
        const employeesLinkedToUser = employees.filter(emp => emp.supervisor === userToDeleteDisplayName);

        if (employeesLinkedToUser.length > 0) {
            alert(`Cannot delete user '${usernameToDelete}' because ${employeesLinkedToUser.length} employee(s) are still linked to '${userToDeleteDisplayName}'. Please reassign these employees first.`);
            return;
        }

        delete users[usernameToDelete];
        localStorage.setItem('users', JSON.stringify(users));
        renderUserList();
        alert(`User '${usernameToDelete}' deleted.`);
    }
}

// Optional: Toggle visibility of Supervisor Name input based on role selection
newUserRoleSelect.addEventListener('change', () => {
    if (newUserRoleSelect.value === 'admin') {
        newSupervisorNameInput.style.display = 'none';
        newSupervisorNameInput.removeAttribute('required');
    } else {
        newSupervisorNameInput.style.display = 'block';
        newSupervisorNameInput.setAttribute('required', 'required');
    }
});
// Initialize visibility on load
newUserRoleSelect.dispatchEvent(new Event('change'));


// --- Excel Export Function ---

exportExcelButton.addEventListener('click', () => {
    if (currentUser.role !== 'admin') {
        alert("Only administrators can export the calendar.");
        return;
    }
    exportCalendarToExcel();
});

function exportCalendarToExcel() {
    const headerRow1 = ['Area', 'Tech #', 'First', 'Last', 'Supervisor', 'Status'];
    const headerRow2 = ['', '', '', '', '', '']; 

    const datesForExport = [];
    // Calculate the actual start date for the displayed range (Sunday of the week)
    const effectiveStartDate = new Date(startDate);
    effectiveStartDate.setDate(effectiveStartDate.getDate() - effectiveStartDate.getDay());

    const tempDate = new Date(effectiveStartDate);
    
    for (let i = 0; i < DISPLAY_DAYS; i++) {
        const dayOfWeekName = tempDate.toLocaleString('en-US', { weekday: 'short' });
        const monthDay = tempDate.toLocaleString('en-US', { month: 'short', day: 'numeric' });
        
        headerRow1.push(`${monthDay}`); 
        headerRow2.push(`${dayOfWeekName}`); 
        
        datesForExport.push(formatDateYYYYMMDD(tempDate));
        tempDate.setDate(tempDate.getDate() + 1);
    }

    const data = [];
    data.push(headerRow1);
    data.push(headerRow2);

    employees.forEach(emp => {
        const rowData = [
            emp.area || '',
            emp.techNumber,
            emp.firstName,
            emp.lastName,
            emp.supervisor || '',
            emp.status || 'Working'
        ];
        
        datesForExport.forEach(dateString => {
            const entry = calendarData[emp.id] ? calendarData[emp.id][dateString] : 'working';
            let statusForExport = '';

            if (typeof entry === 'object' && entry.status === 'pending') {
                statusForExport = `PENDING (${entry.requestedStatus.toUpperCase()})`;
            } else if (typeof entry === 'string') {
                statusForExport = entry.toUpperCase();
            } else {
                statusForExport = 'WORKING';
            }
            rowData.push(statusForExport);
        });
        data.push(rowData);
    });

    const ws = XLSX.utils.aoa_to_sheet(data);
    
    const wscols = [];
    for (let i = 0; i < headerRow1.length; i++) {
        wscols.push({wch: 10});
        if (i < 5) {
            wscols[i].wch = 15;
        }
    }
    ws['!cols'] = wscols;

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Employee Calendar");

    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    saveAs(new Blob([wbout], { type: 'application/octet-stream' }), 'employee_calendar.xlsx');
    alert("Calendar exported to Excel!");
}


// --- Initialization ---
checkAuth();
startDate = new Date(); // Re-initialize startDate to today's date on load
renderShiftsCalendar();
