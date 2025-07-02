/* global Office */

Office.onReady((info) => {
    if (info.host === Office.HostApplication.Outlook) {
        document.addEventListener("DOMContentLoaded", function() {
            initializePlugin();
        });
    }
});

// Configuration - Your actual ESBAAR credentials
const CONFIG = {
    BIGIN_API_BASE: 'https://www.zohoapis.com/bigin/v1',
    POWER_AUTOMATE_WEBHOOK: 'https://prod-150.westeurope.logic.azure.com:443/workflows/c5e04bfa059a48c1b702353fd98058c0/triggers/manual/paths/invoke?api-version=2016-06-01',
    ACCESS_TOKEN: '1000.c2a8a661822a16147f6c8eae88ee0da.b11893cac7a9d74a59c902046bf1bbd0'
};

// Your actual ESBAAR engineers
const MOCK_ENGINEERS = {
    'a.alhinaai@esbaar.com': { name: 'Ali Al Hinaai' },
    'ilyas@esbaar.com': { name: 'Ilyas Said Al Mubsali' },
    'b.shabbir@esbaar.com': { name: 'Burhan Shabbir' },
    'p.kumar@esbaar.com': { name: 'Pradeesh Kumar' }
};

// Mock data for customer auto-completion (you can expand this)
const MOCK_CUSTOMERS = {
    'john@abc.com': {
        name: 'John Smith',
        company: 'ABC Corporation',
        address: '123 Business St, New York, NY 10001'
    },
    'mary@techcorp.com': {
        name: 'Mary Johnson',
        company: 'TechCorp Solutions',
        address: '456 Innovation Ave, San Francisco, CA 94105'
    }
};

function initializePlugin() {
    console.log('ESBAAR Bigin Lead Creator Plugin Initialized');
    
    // Auto-populate email data when plugin loads
    populateEmailData();
    
    // Set up event listeners
    setupEventListeners();
}

function setupEventListeners() {
    // Customer email auto-completion
    document.getElementById('customerEmail').addEventListener('blur', handleCustomerEmailChange);
    
    // Engineer email auto-completion
    document.getElementById('engineerEmail').addEventListener('blur', handleEngineerEmailChange);
    
    // Form submission
    document.getElementById('leadForm').addEventListener('submit', handleFormSubmit);
    
    // Clear form button
    document.getElementById('clearForm').addEventListener('click', clearForm);
}

function populateEmailData() {
    Office.context.mailbox.item.dateTimeReceived.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const receivedDate = new Date(result.value);
            
            // Format for datetime-local input
            const formattedDate = receivedDate.toISOString().slice(0, 16);
            document.getElementById('leadDate').value = formattedDate;
        }
    });

    // Get sender information
    Office.context.mailbox.item.from.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const sender = result.value;
            document.getElementById('customerEmail').value = sender.emailAddress;
            
            // Trigger auto-completion
            handleCustomerEmailChange();
        }
    });
}

function handleCustomerEmailChange() {
    const email = document.getElementById('customerEmail').value.toLowerCase();
    
    if (MOCK_CUSTOMERS[email]) {
        const customer = MOCK_CUSTOMERS[email];
        document.getElementById('customerName').value = customer.name;
        document.getElementById('companyName').value = customer.company;
        document.getElementById('companyAddress').value = customer.address;
    } else {
        // Clear fields if no match found
        document.getElementById('customerName').value = '';
        document.getElementById('companyName').value = '';
        document.getElementById('companyAddress').value = '';
        
        // Try to extract name from email (basic extraction)
        if (email) {
            const nameGuess = email.split('@')[0].replace(/[._]/g, ' ');
            document.getElementById('customerName').value = nameGuess;
        }
    }
}

function handleEngineerEmailChange() {
    const email = document.getElementById('engineerEmail').value.toLowerCase();
    
    if (MOCK_ENGINEERS[email]) {
        document.getElementById('engineerName').value = MOCK_ENGINEERS[email].name;
        
        // Auto-assign to same engineer if not already set
        if (!document.getElementById('assignedTo').value) {
            document.getElementById('assignedTo').value = email;
        }
    } else {
        document.getElementById('engineerName').value = '';
    }
}

async function handleFormSubmit(event) {
    event.preventDefault();
    
    showLoading(true);
    hideMessages();
    
    try {
        // Collect form data
        const formData = collectFormData();