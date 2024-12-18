let ticketsList = [];

Office.onReady(info => {
    if (info.host === Office.HostType.Outlook) {
        Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, handleItemChanged);
        initializeTaskPane();
    }

    document.getElementById("loginButton").onclick = loginToHaloPSA;
    document.getElementById("importButton").onclick = importEmailAsNote;
    document.getElementById("logoutButton").onclick = logout;

    const loginInputs = document.querySelectorAll("#username, #password");
    loginInputs.forEach(input => {
        input.addEventListener("keydown", function(event) {
            if (event.key === "Enter") {
                event.preventDefault();
                loginToHaloPSA();
            }
        });
    });

    const ticketSearchInput = document.getElementById('ticketSearch');
    ticketSearchInput.addEventListener('input', handleTicketSearch);

    document.getElementById("darkModeToggle").addEventListener("click", toggleDarkMode);
});

function initializeTaskPane() {
    const accessToken = sessionStorage.getItem('accessToken');
    const refreshToken = localStorage.getItem('refreshToken');

    if (accessToken) {
        showMainScreen();
        loadTickets();
    } else if (refreshToken) {
        refreshAccessToken()
        .then(() => {
            showMainScreen();
            loadTickets();
        })
        .catch(() => {
            showLoginScreen();
        });
    } else {
        showLoginScreen();
    }
}

function showLoginScreen() {
    document.getElementById('loginScreen').style.display = 'block';
    document.getElementById('mainScreen').style.display = 'none';
    document.getElementById('logoutButton').style.display = 'none';
}

function showMainScreen() {
    document.getElementById('loginScreen').style.display = 'none';
    document.getElementById('mainScreen').style.display = 'block';
    document.getElementById('logoutButton').style.display = 'block';
}

function loginToHaloPSA() {
    const username = document.getElementById('username').value.trim();
    const password = document.getElementById('password').value.trim();

    if (!username || !password) {
        displayStatusMessage('loginError', 'Please enter both username and password.', 'error');
        return;
    }

    showSpinner('loginLoading');

    const loginData = {
        grant_type: 'password',
        client_id: 'YOUR-CLIENTID',
        username: username,
        password: password,
        scope: 'all offline_access'
    };

    fetch('https://YOUR-TENANT.halopsa.com/auth/token', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
        },
        body: new URLSearchParams(loginData)
    })
    .then(response => {
        hideSpinner('loginLoading');

        if (!response.ok) {
            return response.text().then(text => { throw new Error(text); });
        }
        return response.json();
    })
    .then(data => {
        sessionStorage.setItem('accessToken', data.access_token);
        localStorage.setItem('refreshToken', data.refresh_token);
        localStorage.setItem('username', username);
        showMainScreen();
        loadTickets();
    })
    .catch(error => {
        console.error('Login failed:', error);
        displayStatusMessage('loginError', 'Login failed. Please check your credentials.', 'error');
    });
}

function refreshAccessToken() {
    const refreshToken = localStorage.getItem('refreshToken');

    if (!refreshToken) {
        console.error('No refresh token available');
        displayStatusMessage('errorMessage', 'Session expired. Please log in again.', 'error');
        return Promise.reject('No refresh token available');
    }

    const refreshData = {
        grant_type: 'refresh_token',
        client_id: 'YOUR-CLIENTID',
        refresh_token: refreshToken
    };

    return fetch('https://YOUR-TENANT.halopsa.com/auth/token', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
        },
        body: new URLSearchParams(refreshData)
    })
    .then(response => {
        if (!response.ok) {
            return response.text().then(text => { throw new Error(text); });
        }
        return response.json();
    })
    .then(data => {
        sessionStorage.setItem('accessToken', data.access_token);
        if (data.refresh_token) {
            localStorage.setItem('refreshToken', data.refresh_token);
        }
        console.log('Access token refreshed');
        return data.access_token;
    })
    .catch(error => {
        console.error('Failed to refresh access token:', error);
        displayStatusMessage('errorMessage', 'Session expired. Please log in again.', 'error');
        return Promise.reject(error);
    });
}

function logout() {
    sessionStorage.clear();
    localStorage.removeItem('refreshToken');
    localStorage.removeItem('username');
    showLoginScreen();
}

function loadTickets() {
    refreshAccessToken().then(accessToken => {
        const username = localStorage.getItem('username');
        showSpinner('loading');

        fetch(`https://YOUR-TENANT.halopsa.com/api/tickets?open_only=true`, {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${accessToken}`
            }
        })
        .then(response => {
            hideSpinner('loading');

            if (!response.ok) {
                return response.text().then(text => { throw new Error(text); });
            }
            return response.json();
        })
        .then(data => {
            console.log('Response Data:', data);

            if (data && Array.isArray(data.tickets)) {
                ticketsList = data.tickets;
                fetchAgentNamesForTickets(ticketsList, accessToken);
            } else {
                console.error('Unexpected data format:', data);
                displayStatusMessage('errorMessage', 'Failed to load tickets. Unexpected data format.', 'error');
            }
        })
        .catch(error => {
            console.error('Error fetching tickets:', error);
            displayStatusMessage('errorMessage', 'Failed to load tickets.', 'error');
        });
    });
}

function fetchAgentNamesForTickets(tickets, accessToken) {
    const agentIds = new Set(tickets.map(ticket => ticket.agent_id).filter(id => id));
    const agentPromises = Array.from(agentIds).map(agentId => {
        return fetch(`https://YOUR-TENANT.halopsa.com/api/agent/${agentId}`, {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${accessToken}`
            }
        })
        .then(response => {
            if (!response.ok) {
                return response.text().then(text => { throw new Error(text); });
            }
            return response.json();
        })
        .then(agentData => ({
            id: agentId,
            name: agentData.name
        }))
        .catch(error => {
            console.error(`Failed to fetch agent ${agentId}:`, error);
            return {
                id: agentId,
                name: 'Unknown'
            };
        });
    });

    Promise.all(agentPromises)
        .then(agentInfos => {
            const agentMap = new Map(agentInfos.map(agentInfo => [agentInfo.id, agentInfo.name]));
            ticketsList.forEach(ticket => {
                ticket.agent_name = agentMap.get(ticket.agent_id) || 'Unknown';
            });
            populateDropdown(ticketsList);
        })
        .catch(error => {
            console.error('Error fetching agent names:', error);
            displayStatusMessage('errorMessage', 'Failed to fetch agent names.', 'error');
            populateDropdown(ticketsList);
        });
}

function populateDropdown(tickets) {
    const searchInput = document.getElementById('ticketSearch');
    const dropdown = document.getElementById('ticketDropdown');

    function filterTickets() {
        const searchText = searchInput.value.toLowerCase();
        const filteredTickets = tickets.filter(ticket => {
            const ticketId = ticket.id ? String(ticket.id).toLowerCase() : '';
            const ticketSummary = ticket.summary ? ticket.summary.toLowerCase() : '';
            return ticketId.includes(searchText) || ticketSummary.includes(searchText);
        });

        renderDropdown(filteredTickets);
    }

    searchInput.addEventListener('input', filterTickets);

    renderDropdown(tickets);
}

function renderDropdown(tickets) {
    const dropdown = document.getElementById('ticketDropdown');
    dropdown.innerHTML = '';

    tickets.forEach(ticket => {
        const option = document.createElement('div');
        option.classList.add('dropdown-item');
        option.dataset.value = ticket.id;

        const ticketText = document.createElement('div');
        ticketText.textContent = `${ticket.id}: ${ticket.summary}`;

        const agentName = document.createElement('div');
        agentName.textContent = `Klant: ${ticket.user_name} - Behandelaar: ${ticket.agent_name}`;
        agentName.style.fontSize = '12px';
        agentName.style.color = '#666';
        option.appendChild(ticketText);
        option.appendChild(agentName);

        dropdown.appendChild(option);

        option.addEventListener('click', function () {
            document.getElementById('ticketSearch').value = `${ticket.id}: ${ticket.summary}`;
            document.getElementById('ticketSearch').setAttribute('data-selected-id', ticket.id);
            dropdown.style.display = 'none';
        });
    });

    dropdown.style.display = tickets.length ? 'block' : 'none';
}

function handleTicketSearch(event) {
    const searchTerm = event.target.value.trim().toLowerCase();
    const filteredTickets = ticketsList.filter(ticket =>
        String(ticket.id).toLowerCase().includes(searchTerm) ||
        ticket.summary.toLowerCase().includes(searchTerm)
    );
    populateDropdown(filteredTickets);
}

function importEmailAsNote() {
    const selectedTicketId = document.getElementById('ticketSearch').getAttribute('data-selected-id');

    if (!selectedTicketId) {
        alert('Please select a ticket first.');
        return;
    }

    Office.context.mailbox.item.body.getAsync("html", function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            let emailBody = result.value;
            emailBody = cleanUpHtml(emailBody);
            processInlineImages(selectedTicketId, emailBody);
        } else {
            console.error("Failed to get email body:", result.error);
        }
    });
}

function processInlineImages(selectedTicketId, emailBody) {
    const attachments = Office.context.mailbox.item.attachments || [];
    const inlineAttachments = attachments.filter(att => att.isInline);
    const regularAttachments = attachments.filter(att => !att.isInline);
    const allPromises = [];

    function getDataUri(type, base64Content) {
        let prefix = '';
        if (type.startsWith('image/')) {
            prefix = `data:${type};base64,`;
        } else if (type === 'application/pdf') {
            prefix = 'data:application/pdf;base64,';
        } else if (type === 'application/msword' || type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
            prefix = 'data:application/msword;base64,';
        } else if (type === 'application/vnd.ms-excel' || type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
            prefix = 'data:application/vnd.ms-excel;base64,';
        } else {
            prefix = `data:${type};base64,`;
        }
        return prefix + base64Content;
    }

    inlineAttachments.forEach(attachment => {
        allPromises.push(
            new Promise((resolve, reject) => {
                Office.context.mailbox.item.getAttachmentContentAsync(attachment.id, function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        const base64Image = result.value.content;
                        const imageName = attachment.name;
                        const dataUri = getDataUri(attachment.contentType, base64Image);
                        const imgTag = `<img src="${dataUri}" alt="${imageName}" />`;
                        const cidPattern = new RegExp(`<img[^>]+cid:${imageName.replace(/([.*+?^=!:${}()|\[\]\/\\])/g, "\\$1")}[^>]*>`, 'gi');
                        emailBody = emailBody.replace(cidPattern, imgTag);
                        resolve();
                    } else {
                        reject(result.error);
                    }
                });
            })
        );
    });

    regularAttachments.forEach(attachment => {
        allPromises.push(
            new Promise((resolve, reject) => {
                Office.context.mailbox.item.getAttachmentContentAsync(attachment.id, function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        const base64Content = result.value.content;
                        const dataUri = getDataUri(attachment.contentType, base64Content);
                        resolve({
                            filename: attachment.name,
                            data_base64: dataUri
                        });
                    } else {
                        reject(result.error);
                    }
                });
            })
        );
    });

    Promise.all(allPromises)
        .then(results => {
            const attachmentData = results.filter(result => result && result.filename);
            sendHtmlToServer(selectedTicketId, emailBody, attachmentData);
        })
        .catch(error => {
            console.error("Error processing attachments:", error);
            displayStatusMessage('errorMessage', 'Failed to process attachments.', 'error');
        });
}

function cleanUpHtml(html) {
    return html.replace(/\s+/g, ' ').replace(/>\s+</g, '><').trim();
}

function sendHtmlToServer(ticketId, htmlContent, attachmentData = []) {
    refreshAccessToken().then(accessToken => {
        showSpinner('loading');

        const payload = [{
            ticket_id: ticketId,
            outcome: "Interne Notitie",
            note_html: htmlContent,
            hiddenfromuser: true,
            attachments: attachmentData
        }];

        fetch('https://YOUR-TENANT.halopsa.com/api/actions', {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(payload)
        })
        .then(response => {
            hideSpinner('loading');
            if (!response.ok) {
                return response.text().then(text => { throw new Error(text); });
            }
            return response.json();
        })
        .then(data => {
            console.log('Response Data:', data);
            displayStatusMessage('successMessage', 'Email successfully imported as a note.', 'success');
        })
        .catch(error => {
            console.error('Error sending email to server:', error);
            displayStatusMessage('errorMessage', 'Failed to import email as a note.', 'error');
        });
    });
}

function showSpinner(id) {
    const spinner = document.getElementById(id);
    if (spinner) {
        spinner.style.display = 'inline-block';
    }
}

function hideSpinner(id) {
    const spinner = document.getElementById(id);
    if (spinner) {
        spinner.style.display = 'none';
    }
}

function displayStatusMessage(id, message, type) {
    const element = document.getElementById(id);
    if (element) {
        element.textContent = message;
        element.style.display = 'block';
        element.style.color = type === 'error' ? 'red' : 'green';
        setTimeout(() => {
            element.style.display = 'none';
        }, 5000);
    } else {
        console.warn(`Element with ID "${id}" not found.`);
    }
}

function handleItemChanged(eventArgs) {
    console.log('Item changed:', eventArgs);
    loadTickets();
}

function toggleDarkMode() {
    const isDarkMode = document.body.getAttribute("data-theme") === "dark";
    document.body.setAttribute("data-theme", isDarkMode ? "light" : "dark");
    localStorage.setItem("theme", isDarkMode ? "light" : "dark");
}

document.addEventListener("DOMContentLoaded", () => {
    const theme = localStorage.getItem("theme") || "light";
    document.body.setAttribute("data-theme", theme);
});
