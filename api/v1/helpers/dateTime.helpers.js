// Importing the moment library for date parsing and manipulation
const moment = require('moment');

// Function to validate the format and relationship of two date strings
function validateDates(startDateString, endDateString) {
    // Parsing the start and end dates using moment.js
    const startDate = moment(startDateString, 'YYYY-MM-DD', true);
    const endDate = moment(endDateString, 'YYYY-MM-DD', true);

    // Checking if the parsed dates are valid and in the specified format
    if (!startDate.isValid() || !endDate.isValid()) {
        throw new Error('Invalid date format. Please use YYYY-MM-DD format.');
    }

    // Checking if the start date is after the end date
    if (startDate.isAfter(endDate)) {
        throw new Error('Start date cannot be after end date.');
    }

    // Additional validation logic can be added here if needed

    // Returning the validated start and end dates
    return { startDate, endDate };
}

// Function to generate the SQL WHERE clause for a date range
function generateDateClause(startDate, endDate) {
    // Creating the WHERE clause string for a date range
    const whereQuery = ` BETWEEN '${startDate}' AND '${endDate}' `;
    return whereQuery;
}

function formatDateTime(date) {
    const days = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31"];
    const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

    const day = days[date.getDate() - 1];
    const month = months[date.getMonth()];
    const year = date.getFullYear();
    const hours = date.getHours().toString().padStart(2, '0');
    const minutes = date.getMinutes().toString().padStart(2, '0');

    return `${day} ${month}, ${year} ${hours}:${minutes}`;
}

// Exporting the functions to be used in other modules
module.exports = {
    validateDates,
    generateDateClause,
    formatDateTime
};
