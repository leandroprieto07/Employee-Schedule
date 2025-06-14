// --- Importaciones de Firebase SDK ---
import { initializeApp } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-app.js";
import { getAuth, signInAnonymously, signInWithCustomToken, onAuthStateChanged, signOut } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-auth.js";
import { getFirestore, collection, doc, setDoc, updateDoc, deleteDoc, onSnapshot, query, where, getDocs } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-firestore.js";

// --- Configuración de Firebase (TU PROPIA CLAVE API REAL AQUÍ) ---
// Obtén esta configuración de tu consola de Firebase -> Configuración del proyecto -> Tus apps -> Configuración de la aplicación web.
const firebaseConfig = {
    // IMPORTANTE: REEMPLAZA ESTE VALOR CON TU CLAVE API REAL DE FIREBASE.
    // La que empieza con "AIzaSy..." y que encontrarás en la configuración de tu proyecto web en Firebase.
    apiKey: "AIzaSyC3u4FV6h1dzqQuWo2U8F7SxCe2AxoiG5M", 
    authDomain: "employee-shift-calendar-81868.firebaseapp.com",
    projectId: "employee-shift-calendar-81868",
    storageBucket: "employee-shift-calendar-81868.firebasestorage.app",
    messagingSenderId: "1000655634680",
    appId: "1:1000655634680:web:de80fa37b44bd9272efb0d",
    measurementId: "G-2NXDB1ZC41"
};

// --- Inicialización de Firebase ---
const firebaseApp = initializeApp(firebaseConfig);
const auth = getAuth(firebaseApp);
const db = getFirestore(firebaseApp);

// Usamos el projectId como el ID de la aplicación para las rutas de Firestore
const appId = firebaseConfig.projectId;

// Firestore path constants
const COLLECTIONS = {
    APP_USERS: 'appUsers',    // Stores user login details (username, password, role, displayName)
    EMPLOYEES: 'employees'    // Stores employee data and their calendar shifts
};

// --- Application State Variables ---
let currentUser = null; // Currently logged-in app user (not Firebase Auth user)
let employees = [];     // Array of employee objects, populated from Firestore
let appUsers = {};      // Map of app users (username -> {password, role, displayName}), populated from Firestore
let firebaseUserId = null; // The actual Firebase Authentication UID


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

const adminMainSection = document.getElementById('admin-main-section');
const employeeManagementSection = document.getElementById('employee-management-section');
const addEmployeeForm = document.getElementById('add-employee-form');
const employeeListBody = document.getElementById('employee-list-body');

const userManagementSection = document.getElementById('user-management-section');
const addUserForm = document.getElementById('add-user-form');
const newUsernameInput = document.getElementById('new-username');
const newPasswordInput = document.getElementById('new-password');
const newSupervisorNameInput = document.getElementById('new-supervisor-name');
const newUserRoleSelect = document.getElementById('new-user-role');
const userCreationMessage = document.getElementById('user-creation-message');
const userListBody = document.getElementById('user-list-body');

const statusModal = document.getElementById('status-modal');
const modalEmployeeName = document.getElementById('modal-employee-name');
const modalDate = document.getElementById('modal-date');
const statusSelect = document.getElementById('status-select');
const saveStatusButton = document.getElementById('save-status-button');
const closeButton = statusModal.querySelector('.close-button');

// --- Calendar Display Variables ---
let startDate = new Date();
const DISPLAY_DAYS = 14;

// Modal editing state
let currentEditingEmployeeId = null;
let currentEditingDate = null;
let currentEditingCell = null;


// --- Utility Functions ---

function formatDateYYYYMMDD(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function formatDateMMDDYYYY(dateStringYYYYMMDD) {
    const [year, month, day] = dateStringYYYYMMDD.split('-');
    return `${month}/${day}/${year}`;
}

function isWeekend(date) {
    const dayOfWeek = date.getDay(); // 0 = Sunday, 6 = Saturday
    return dayOfWeek === 0 || dayOfWeek === 6;
}


// --- Firebase Authentication Management ---

// This function handles the initial Firebase Auth state and sets up Firestore listeners.
onAuthStateChanged(auth, async (user) => {
    if (user) {
        firebaseUserId = user.uid;
        console.log("Firebase Authenticated User ID:", firebaseUserId);
        // Setup Firestore listeners once authenticated
        setupFirestoreListeners();
        // Try to log in with previously saved currentUser from localStorage
        if (localStorage.getItem('currentUser')) {
             currentUser = JSON.parse(localStorage.getItem('currentUser'));
             // Ensure displayName is updated from appUsers for current session
             if (currentUser.role === 'supervisor' && appUsers[currentUser.username] && appUsers[currentUser.username].displayName) {
                currentUser.displayName = appUsers[currentUser.username].displayName;
             } else if (currentUser.role === 'supervisor') {
                currentUser.displayName = currentUser.username;
             }
             renderApp();
        } else {
            // If no app user was previously selected, show the login form
            loginContainer.style.display = 'block';
            appContainer.style.display = 'none';
        }
    } else {
        // No Firebase user authenticated, try to sign in anonymously
        try {
            await signInAnonymously(auth);
            console.log("Signed in anonymously to Firebase.");
        } catch (error) {
            console.error("Firebase Auth error during anonymous sign-in:", error);
            loginMessage.textContent = "Authentication error. Please refresh.";
            loginContainer.style.display = 'block';
            appContainer.style.display = 'none';
        }
    }
});


// --- Firestore Data Listeners ---

function setupFirestoreListeners() {
    if (!db || !appId || !firebaseUserId) {
        console.warn("Firestore not initialized or user not authenticated for listeners.");
        return;
    }

    // Listener for App Users
    onSnapshot(collection(db, `artifacts/${appId}/public/data/${COLLECTIONS.APP_USERS}`), (snapshot) => {
        const fetchedUsers = {};
        snapshot.forEach(doc => {
            fetchedUsers[doc.id] = doc.data();
        });
        appUsers = fetchedUsers;
        console.log("Fetched App Users:", appUsers);

        // If appUsers are empty, create default admin/supervisor
        if (Object.keys(appUsers).length === 0) {
            console.log("No app users found. Creating default users...");
            createDefaultAppUsers();
        }

        // Re-render user list and potentially re-authenticate current user's role if appUsers updated
        if (currentUser && appUsers[currentUser.username]) {
            currentUser.role = appUsers[currentUser.username].role;
            currentUser.displayName = appUsers[currentUser.username].displayName || currentUser.username;
            localStorage.setItem('currentUser', JSON.stringify(currentUser)); // Update local storage
        }
        renderUserList(); // Ensure user list is up-to-date
        renderApp(); // Re-render app to update UI based on new user roles/permissions
    }, (error) => {
        console.error("Error fetching app users:", error);
    });

    // Listener for Employees (and their calendar data)
    onSnapshot(collection(db, `artifacts/${appId}/public/data/${COLLECTIONS.EMPLOYEES}`), (snapshot) => {
        employees = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
        console.log("Fetched Employees:", employees);
        renderEmployeeList(); // Re-render employee list
        renderShiftsCalendar(); // Re-render calendar
    }, (error) => {
        console.error("Error fetching employees:", error);
    });
}

// Function to create default admin and supervisor users in Firestore
async function createDefaultAppUsers() {
    if (!db) {
        console.error("Firestore DB not initialized.");
        return;
    }
    try {
        await setDoc(doc(db, `artifacts/${appId}/public/data/${COLLECTIONS.APP_USERS}`, 'admin'), {
            password: 'adminpassword',
            role: 'admin',
            displayName: 'Admin User'
        });
        await setDoc(doc(db, `artifacts/${appId}/public/data/${COLLECTIONS.APP_USERS}`, 'supervisor1'), {
            password: 'sup1password',
            role: 'supervisor',
            displayName: 'Supervisor Alpha'
        });
        console.log("Default users created in Firestore.");
    } catch (error) {
        console.error("Error creating default users in Firestore:", error);
    }
}


// --- Login/Logout Functions (now linked to appUsers from Firestore) ---

loginForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    const username = document.getElementById('username').value;
    const password = document.getElementById('password').value;

    const userInDb = appUsers[username]; // Check against Firestore-loaded appUsers

    if (userInDb && userInDb.password === password) {
        currentUser = { username: username, role: userInDb.role, displayName: userInDb.displayName || username };
        localStorage.setItem('currentUser', JSON.stringify(currentUser)); // Save selected app user locally
        loginMessage.textContent = '';
        renderApp();
    } else {
        loginMessage.textContent = 'Incorrect username or password.';
    }
});

logoutButton.addEventListener('click', async () => {
    currentUser = null;
    localStorage.removeItem('currentUser'); // Clear selected app user
    // No Firebase signOut here, as authentication is handled by Canvas environment or anonymous sign-in
    loginContainer.style.display = 'block';
    appContainer.style.display = 'none';
});


/**
 * Renders the main application interface after a successful login.
 * Shows/hides sections based on the user's role.
 */
function renderApp() {
    if (!currentUser) { // If no app user is selected yet
        loginContainer.style.display = 'block';
        appContainer.style.display = 'none';
        return;
    }

    loginContainer.style.display = 'none';
    appContainer.style.display = 'block';
    const displayUserName = currentUser.role === 'supervisor' && currentUser.displayName ? currentUser.displayName : currentUser.username;
    userInfoSpan.textContent = `Welcome, ${displayUserName} (${currentUser.role})`;

    if (currentUser.role === 'admin') {
        adminMainSection.style.display = 'block';
        employeeManagementSection.style.display = 'block';
        userManagementSection.style.display = 'block';
    } else {
        adminMainSection.style.display = 'none';
        employeeManagementSection.style.display = 'none';
        userManagementSection.style.display = 'none';
    }

    renderShiftsCalendar();
    renderEmployeeList();
    renderUserList();
}

// --- Horizontal Shift Calendar Functions ---

/**
 * Renders the horizontal shift calendar table.
 * Dynamically generates date headers and employee rows.
 */
function renderShiftsCalendar() {
    dateHeaderRow.innerHTML = '';
    shiftsTableBody.innerHTML = '';

    // Adjust startDate to be the beginning of the week (Sunday)
    const displayStartDate = new Date(startDate); 
    displayStartDate.setDate(displayStartDate.getDate() - displayStartDate.getDay());

    const currentDay = new Date(displayStartDate);
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

        // Employee's calendar data is now part of the employee document in Firestore
        // Use employee.calendar map, defaulting to empty object if not exists
        const employeeCalendar = employee.calendar || {}; 
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
            if (currentUser) { // Ensure currentUser is not null
                if (currentUser.role === 'admin' || 
                   (currentUser.role === 'supervisor' && employee.supervisor === currentUser.displayName)) {
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

prevWeekButton.addEventListener('click', () => {
    startDate.setDate(startDate.getDate() - DISPLAY_DAYS);
    renderShiftsCalendar();
});

nextWeekButton.addEventListener('click', () => {
    startDate.setDate(startDate.getDate() + DISPLAY_DAYS);
    renderShiftsCalendar();
});

// --- Day Status Modal ---

async function openStatusModal(employee, dateString, currentEntry, cell) {
    if (!currentUser) { // Ensure user is logged into the app before opening modal
        alert("Please log in to manage shifts.");
        return;
    }
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
        
        // Remove previous event listeners to prevent duplicates
        const approveBtn = document.getElementById('approve-request-button');
        const rejectBtn = document.getElementById('reject-request-button');
        
        // Clone and replace to remove old event listeners
        const newApproveBtn = approveBtn.cloneNode(true);
        const newRejectBtn = rejectBtn.cloneNode(true);
        approveBtn.parentNode.replaceChild(newApproveBtn, approveBtn);
        rejectBtn.parentNode.replaceChild(newRejectBtn, rejectBtn);

        newApproveBtn.addEventListener('click', async () => {
            const employeeRef = doc(db, `artifacts/${appId}/public/data/${COLLECTIONS.EMPLOYEES}`, currentEditingEmployeeId);
            try {
                await updateDoc(employeeRef, {
                    [`calendar.${currentEditingDate}`]: currentEntry.requestedStatus
                });
                console.log("Request approved and status updated in Firestore.");
                statusModal.style.display = 'none';
                alert("Request approved.");
            } catch (error) {
                console.error("Error approving request:", error);
                alert("Error approving request. Check console.");
            }
        });

        newRejectBtn.addEventListener('click', async () => {
            const employeeRef = doc(db, `artifacts/${appId}/public/data/${COLLECTIONS.EMPLOYEES}`, currentEditingEmployeeId);
            try {
                await updateDoc(employeeRef, {
                    [`calendar.${currentEditingDate}`]: 'working' // Revert to 'working'
                });
                console.log("Request rejected and status reverted in Firestore.");
                statusModal.style.display = 'none';
                alert("Request rejected.");
            } catch (error) {
                console.error("Error rejecting request:", error);
                alert("Error rejecting request. Check console.");
            }
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

saveStatusButton.addEventListener('click', async () => {
    const newStatus = statusSelect.value;
    const employeeToEdit = employees.find(emp => emp.id === currentEditingEmployeeId);

    if (!employeeToEdit) {
        alert("Employee not found.");
        return;
    }

    // Re-check supervisor linkage for robustness (using displayName)
    if (currentUser.role === 'supervisor' && employeeToEdit.supervisor !== currentUser.displayName) {
        alert("You can only manage shifts for your assigned employees.");
        statusModal.style.display = 'none';
        return;
    }

    const employeeRef = doc(db, `artifacts/${appId}/public/data/${COLLECTIONS.EMPLOYEES}`, currentEditingEmployeeId);

    try {
        if (currentUser.role === 'supervisor') {
            await updateDoc(employeeRef, {
                [`calendar.${currentEditingDate}`]: {
                    status: 'pending',
                    requestedStatus: newStatus,
                    requestedBy: currentUser.displayName,
                }
            });
            statusModal.style.display = 'none';
            alert(`Request for ${newStatus.toUpperCase()} sent for approval.`);

        } else if (currentUser.role === 'admin') {
            // Admin makes a direct change (overriding any pending status)
            await updateDoc(employeeRef, {
                [`calendar.${currentEditingDate}`]: newStatus
            });
            statusModal.style.display = 'none';
            alert("Status updated directly by Admin.");
        } else {
            alert("You do not have permissions to modify the status.");
            statusModal.style.display = 'none';
        }
    } catch (error) {
        console.error("Error updating status:", error);
        alert("Error updating status. Check console.");
    }
});


// --- Employee Management Functions (Admin) ---

addEmployeeForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    if (!currentUser || currentUser.role !== 'admin') {
        alert("Only administrators can add employees.");
        return;
    }

    const area = document.getElementById('employee-area').value;
    const techNumber = document.getElementById('employee-tech-number').value;
    const firstName = document.getElementById('employee-first-name').value;
    const lastName = document.getElementById('employee-last-name').value;
    const supervisor = document.getElementById('employee-supervisor').value.trim(); // Supervisor's DISPLAY name

    // Validate that the entered supervisor display name exists as a supervisor
    const supervisorExists = Object.values(appUsers).some(user => 
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
        area,
        techNumber,
        firstName,
        lastName,
        supervisor,
        calendar: {} // Initialize with an empty calendar map for Firestore
    };

    try {
        // Firestore automatically generates document ID if not provided
        await setDoc(doc(collection(db, `artifacts/${appId}/public/data/${COLLECTIONS.EMPLOYEES}`)), newEmployee);
        addEmployeeForm.reset();
        alert("Employee added successfully!");
    } catch (error) {
        console.error("Error adding employee:", error);
        alert("Error adding employee. Check console.");
    }
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
                ${currentUser && currentUser.role === 'admin' ? `<button onclick="editEmployee('${emp.id}')">Edit</button>
                                                   <button onclick="deleteEmployee('${emp.id}')">Delete</button>` : ''}
            </td>
        `;
        employeeListBody.appendChild(row);
    });
}

async function editEmployee(id) {
    if (!currentUser || currentUser.role !== 'admin') {
        alert("Only administrators can edit employees.");
        return;
    }
    const employee = employees.find(emp => emp.id === id);
    if (employee) {
        const newArea = prompt("New Area:", employee.area);
        const newTechNumber = prompt("New Tech #:", employee.techNumber);
        const newFirstName = prompt("New First Name:", employee.firstName);
        const newLastName = prompt("New Last Name:", employee.lastName);
        const newSupervisor = prompt("New Supervisor (Display Name):", employee.supervisor);
        
        if (newArea === null || newTechNumber === null || newFirstName === null || newLastName === null || newSupervisor === null) {
            // User cancelled one of the prompts
            return;
        }

        const trimmedSupervisor = newSupervisor.trim();

        // Validate that the entered supervisor display name exists as a supervisor
        const supervisorExists = Object.values(appUsers).some(user => 
            user.role === 'supervisor' && user.displayName === trimmedSupervisor
        );

        if (!supervisorExists && trimmedSupervisor !== '') {
            alert("The linked Supervisor Display Name does not exist or is not a supervisor user. Please enter an existing supervisor's display name.");
            return; // Stop editing if invalid supervisor
        }

        // Check if Tech # already exists for another employee (excluding the current one being edited)
        if (employees.some(emp => emp.id !== id && emp.techNumber === newTechNumber)) {
            alert("An employee with this Tech # already exists.");
            return;
        }

        const employeeRef = doc(db, `artifacts/${appId}/public/data/${COLLECTIONS.EMPLOYEES}`, id);
        try {
            await updateDoc(employeeRef, {
                area: newArea,
                techNumber: newTechNumber,
                firstName: newFirstName,
                lastName: newLastName,
                supervisor: trimmedSupervisor
            });
            alert("Employee updated successfully!");
        } catch (error) {
            console.error("Error updating employee:", error);
            alert("Error updating employee. Check console.");
        }
    }
}

async function deleteEmployee(id) {
    if (!currentUser || currentUser.role !== 'admin') {
        alert("Only administrators can delete employees.");
        return;
    }
    // Use custom modal for confirm in real app, alert for demo
    if (confirm("Are you sure you want to delete this employee? This cannot be undone.")) {
        const employeeRef = doc(db, `artifacts/${appId}/public/data/${COLLECTIONS.EMPLOYEES}`, id);
        try {
            await deleteDoc(employeeRef);
            alert("Employee deleted successfully!");
        } catch (error) {
            console.error("Error deleting employee:", error);
            alert("Error deleting employee. Check console.");
        }
    }
}

// --- User Management Functions (Admin) ---

addUserForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    if (!currentUser || currentUser.role !== 'admin') {
        alert("Only administrators can manage users.");
        return;
    }

    const newUsername = newUsernameInput.value.trim();
    const newPassword = newPasswordInput.value.trim();
    const newUserRole = newUserRoleSelect.value;
    let newSupervisorDisplayName = newSupervisorNameInput.value.trim();

    userCreationMessage.textContent = '';

    if (!newUsername || !newPassword) {
        userCreationMessage.textContent = 'Username and password cannot be empty.';
        return;
    }

    if (appUsers[newUsername]) { // Check against Firestore-loaded appUsers
        userCreationMessage.textContent = `User '${newUsername}' already exists.`;
        return;
    }

    const newUserDoc = {
        password: newPassword,
        role: newUserRole
    };

    if (newUserRole === 'supervisor') {
        newUserDoc.displayName = newSupervisorDisplayName || newUsername;
    } else {
        delete newUserDoc.displayName; // Ensure displayName is not set for admin users
    }
    
    try {
        // Use username as the document ID for appUsers for easy lookup
        await setDoc(doc(db, `artifacts/${appId}/public/data/${COLLECTIONS.APP_USERS}`, newUsername), newUserDoc);
        addUserForm.reset();
        newSupervisorNameInput.value = '';
        userCreationMessage.textContent = `User '${newUsername}' (${newUserRole}) created successfully.`;
        userCreationMessage.style.color = 'green';
    } catch (error) {
        console.error("Error creating user:", error);
        userCreationMessage.textContent = "Error creating user. Check console.";
        userCreationMessage.style.color = 'red';
    }
});

function renderUserList() {
    userListBody.innerHTML = '';
    for (const username in appUsers) { // Iterate over appUsers from Firestore
        const user = appUsers[username];
        const row = document.createElement('tr');
        const userDisplayInfo = user.role === 'supervisor' && user.displayName ? `(Display: ${user.displayName})` : '';
        row.innerHTML = `
            <td>${username}</td>
            <td>${user.role} ${userDisplayInfo}</td>
            <td>
                ${currentUser && currentUser.role === 'admin' && username !== currentUser.username ? `
                    <button onclick="deleteUser('${username}')" style="background-color: #dc3545;">Delete</button>
                ` : ''}
            </td>
        `;
        userListBody.appendChild(row);
    }
}

async function deleteUser(usernameToDelete) {
    if (!currentUser || currentUser.role !== 'admin') {
        alert("Only administrators can delete users.");
        return;
    }
    if (usernameToDelete === currentUser.username) {
        alert("You cannot delete your own user account.");
        return;
    }

    // Get the display name of the user to delete
    const userToDeleteDisplayName = appUsers[usernameToDelete]?.displayName || usernameToDelete;

    // Check if any employees are linked to this supervisor's displayName
    const q = query(
        collection(db, `artifacts/${appId}/public/data/${COLLECTIONS.EMPLOYEES}`),
        where('supervisor', '==', userToDeleteDisplayName)
    );
    
    try {
        const querySnapshot = await getDocs(q);
        if (querySnapshot.size > 0) {
            alert(`Cannot delete user '${usernameToDelete}' because ${querySnapshot.size} employee(s) are still linked to '${userToDeleteDisplayName}'. Please reassign these employees first.`);
            return;
        }

        // If no linked employees, proceed with deletion
        if (confirm(`Are you sure you want to delete user '${usernameToDelete}'? This cannot be undone.`)) {
            await deleteDoc(doc(db, `artifacts/${appId}/public/data/${COLLECTIONS.APP_USERS}`, usernameToDelete));
            alert(`User '${usernameToDelete}' deleted.`);
        }
    } catch (error) {
        console.error("Error checking/deleting user:", error);
        alert("Error deleting user. Check console.");
    }
}

// Optional: Toggle visibility of Supervisor Name input based on role selection
// This listener needs to be set up after DOM elements are available
document.addEventListener('DOMContentLoaded', () => {
    if (newUserRoleSelect && newSupervisorNameInput) {
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
    }
});


// --- Excel Export Function ---

exportExcelButton.addEventListener('click', () => {
    if (!currentUser || currentUser.role !== 'admin') {
        alert("Only administrators can export the calendar.");
        return;
    }
    exportCalendarToExcel();
});

function exportCalendarToExcel() {
    const headerRow1 = ['Area', 'Tech #', 'First', 'Last', 'Supervisor', 'Status'];
    const headerRow2 = ['', '', '', '', '', '']; 

    const datesForExport = [];
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
            // Get entry from employee.calendar field
            const entry = emp.calendar ? emp.calendar[dateString] : 'working';
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
