// PizZip is required because docx/pptx/xlsx files are all zipped files, and
// the PizZip library allows us to load the file in memory
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const { readFileSync, writeFileSync } = require('fs');
const fs = require('fs');
const htmlToText = require('html-to-text');

/**
 * Check if the file exists at the specified path.
 * @param {string} filePath - Path of the file to check.
 * @returns {boolean} True if the file exists, false otherwise.
 */
function checkFileExists(filePath) {
    try {
        fs.accessSync(filePath, fs.constants.F_OK);
        return true; // File exists
    } catch (err) {
        return false; // File does not exist
    }
}

/**
 * Remove empty objects from an array of objects.
 * @param {Array} jsonArray - Array of objects.
 * @returns {Array} Array with empty objects removed.
 */
function removeEmptyObjects(jsonArray) {
    return jsonArray.filter(obj => Object.keys(obj).length !== 0);
}

/**
 * Convert Chat Conversation JSON Object structure to DOCX format.
 * @param {Array} originalJSON - Original JSON structure.
 * @returns {Array} Converted JSON in render format.
 */
function convertJSONToRender(originalJSON) {
    let convertedJSON = [];
    originalJSON.forEach(obj => {
        let newObj = {};
        if (obj.bot) {
            newObj.text =  obj.bot;
            newObj.head = 'HIVE: ';
        } else if (obj.user) {
            newObj.text = obj.user;
            newObj.head = 'USER: ';
        }
        if (obj.citation){
            if(obj.citation.length > 0){
                console.log('obj.citation =>', obj.citation)
                newObj.citation = obj.citation;
            }
        }
        convertedJSON.push(newObj);
    });
    return convertedJSON;
}

/**
 * Initial Convert input JSON format to a desired output format .
 * @param {Array} inputJSON - Input JSON array.
 * @returns {Array} Formatted JSON array.
 */
function convertFormatJSON(inputJSON) {
    return inputJSON.map(item => {
        const { sender, text, metadata } = item;
        return { [sender]: text, metadata };
    });
}

// Function to remove special characters from a string and replace them with spaces
function removeSpecialCharactersWithSpace(str) {
    // Define regex pattern to match allowed characters: letters, digits, hyphen, and underscore
    //const regex = /[^a-zA-Z0-9\s]/g;

    //const regex = /[^a-zA-Z0-9\s\[\]]/g;

    // Keep the "." and "'" (Single quotes) character to search query string
    //const regex = /[^a-zA-Z0-9\s\[\].'"]/g;

    const regex = /[^a-zA-Z0-9\s\[\].'%"]/g;
    
    // Replace special characters with spaces
    const cleanStr = str.replace(regex, ' ');
    
    // Return the cleaned string
    return cleanStr
}

  // Has the string is only single special characters
  function hasSingleSpecialCharacter(inputString) {
    // Regular expression to match a single special character
    var specialCharacterRegex = /^[^\w\s]*[_!@#$%^&*()-][^\w\s]*$/;

    // Test if the input string matches the regular expression
    return specialCharacterRegex.test(inputString);
}

/**
 * Convert JSON object to a TXT file.
 * @param {Object} jsonObject - The JSON object to convert.
 * @param {string} outputFilePath - The file path for the output TXT file.
 * @returns {Object} Object indicating the status of the conversion.
 */
function convertJSONToTXT(jsonObject, outputFilePath) {
    let txtContent = ''; // Initialize text content

    /**
     * Recursive function to stringify nested objects.
     * @param {Object} obj - The object to stringify.
     * @param {number} depth - The depth of the recursion (default is 0).
     */
    function stringifyObject(obj, depth = 0) {
        const indent = '  '.repeat(depth); // Indentation based on depth

        // Loop through each key-value pair in the object
        for (const [key, value] of Object.entries(obj)) {
            if (typeof value === 'object' && value !== null) {
                // If the value is an object, recursively stringify it
                txtContent += `${indent}${key}:\n`; // Key with indentation
                stringifyObject(value, depth + 1); // Recursive call with increased depth
            } else {
                // If the value is not an object, append it to the text content
                txtContent += `${indent}${key}: ${value}\n`; // Key-value pair with indentation
            }
        }
    }

    // Start the conversion by calling the recursive function
    stringifyObject(jsonObject);

    try {
        // Write the content to a TXT file
        fs.writeFileSync(outputFilePath, txtContent);
        return { status: true }; // Return status indicating success
    } catch (error) {
        console.error('Error writing file:', error);
        return { status: false }; // Return status indicating failure
    }
}

/**
 * Check if obj.metadata is an array containing objects.
 * @param {Object} obj - The object containing the metadata property.
 * @returns {boolean} True if obj.metadata is an array containing objects, false otherwise.
 */
function isArrayOfObjects(obj) {
    console.log('isArrayOfObjects =>', obj)
    // Check if obj.metadata exists and is an array
    if (obj && Array.isArray(obj)) {
      // Check if each element of the array is an object
      return obj && typeof obj === 'object' && Object.keys(obj).length > 0;
    }
    return false;
  }


/**
 * Render multiple objects into a DOCX file using a template.
 * @param {string} templateFilePath - Path to the template file.
 * @param {string} outputFileName - Name of the output DOCX file.
 * @param {Array} data - Array of objects to render.
 * @returns {Object} Object containing status and message.
 */
function renderMultipleObjects(templateFilePath, outputFileName, data) {
    console.log('Calling renderMultipleObjects Function*******************');
    try {
        // Load the template file as binary content
        const content = readFileSync(templateFilePath, 'binary');
        // Unzip the content of the file
        const zip = new PizZip(content);
        // Create a new Docxtemplater instance
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true
        });
        // Set the data for rendering
        doc.setData({ messages: data, saved_Date: new Date() });
        // Render the document
        doc.render();
        // Generate buffer from the rendered document
        const buf = doc.getZip().generate({
            type: 'nodebuffer',
            compression: 'DEFLATE',
        });
        // Write the DOCX file
        writeFileSync(outputFileName, buf);
        console.log('DOCX file created successfully:', outputFileName);
        return { status: true, message: 'DOCX file created successfully:', outputFileName };
    } catch (error) {
        console.error('Error creating DOCX file:', error);
        return { status: false, message: error };
    }
}

function splitBySquareBrackets(str) {
    const result = []; // Array to store the resulting substrings
    let startIndex = 0; // Starting index of the current substring
    let openCount = 0; // Counter to keep track of open square brackets

    // Loop through each character in the string
    for (let i = 0; i < str.length; i++) {
        if (str[i] === '[') { // If an open square bracket is encountered
            if (openCount === 0) { // If this is the first open square bracket encountered
                // Push the substring before the open square bracket to the result array
                result.push(str.substring(startIndex, i));
                // Update the starting index to the character after the open square bracket
                startIndex = i + 1;
            }
            openCount++; // Increment the open bracket count
        } else if (str[i] === ']') { // If a close square bracket is encountered
            openCount--; // Decrement the open bracket count
            if (openCount === 0) { // If this is the matching close square bracket
                // Push the substrings enclosed within the square brackets to the result array
                result.push(str.substring(startIndex, i).split(' ').filter(Boolean));
                // Update the starting index to the character after the close square bracket
                startIndex = i + 1;
            }
        }
    }

    // Push the remaining substring after the last square bracket to the result array
    if (startIndex < str.length) {
        result.push(str.substring(startIndex).split(' ').filter(Boolean));
    }

    // Flatten the result array and return
    return result.flat();
}

// Remove duplicate emails from data array
function removeDuplicateEmails(data) {
    try {
      // Create a Set to store unique email addresses
      const uniqueEmails = new Set();
  
      // Filter the data array to remove duplicate email entries
      const uniqueData = data.filter(item => {
        // Check if the email is already in the Set
        // If not, add it to the Set and return true to keep the item
        if (!uniqueEmails.has(item.user_email)) {
          uniqueEmails.add(item.user_email);
          return true;
        }
        // If the email is already in the Set, return false to remove the item
        return false;
      });
  
      return uniqueData;
    } catch (error) {
      // Handle any errors that occur during filtering
      throw new Error(`Error in removeDuplicateEmails: ${error.message}`);
    }
  }
  
/**
 * Sorts an array of user objects alphabetically based on the 'user_name' field.
 * @param {Array} users - The array of user objects to be sorted.
 * @returns {Array} - The sorted array of user objects.
 */
function sortUsersAlphabetically(users) {
    // Use the Array.sort() method to sort the array of users.
    // The sort function takes two parameters, 'a' and 'b', representing two elements to compare.
    return users.sort((a, b) => {
        // Convert both user names to lowercase for case-insensitive sorting.
        const nameA = a.user_name.toLowerCase();
        const nameB = b.user_name.toLowerCase();

        // Compare the lowercase user names.
        // If nameA comes before nameB, return a negative value to place 'a' before 'b' in the sorted array.
        if (nameA < nameB) {
            return -1;
        }
        // If nameA comes after nameB, return a positive value to place 'a' after 'b' in the sorted array.
        if (nameA > nameB) {
            return 1;
        }
        // If both names are equal, return 0 to maintain their relative order.
        return 0;
    });
}

/**
 * Removes duplicate user emails from the second JSON array based on user emails present in the first JSON array.
 * @param {Array} firstJson - The first JSON array containing objects with user emails.
 * @param {Array} secondJson - The second JSON array containing objects with user emails.
 * @returns {Array} - The filtered second JSON array with duplicate user emails removed.
 */
function removeDuplicatesBasedOnEmail(firstJson, secondJson) {
    // Extract unique user emails from the first JSON array
    const uniqueEmails = new Set();
    firstJson.forEach(item => {
        uniqueEmails.add(item.user_email);
    });

    // Filter the second JSON array to remove objects with user emails present in the first JSON array
    const filteredSecondJson = secondJson.filter(item => {
        return !uniqueEmails.has(item.user_email);
    });

    return filteredSecondJson;
}

// Modified saveEmailsToTxtFile method to handle grouped conversations
function saveEmailsToTxtFile(conversations, filePath) {
    let content = '';

    conversations.forEach(conversation => {
        // Add the subject with a more prominent separator
        content += `===================================================================================================\n`;
        //content += `Conversation ID: ${conversation.conversation_id}\n`;
        content += `Subject: ${conversation.subject}\n`;
        content += `===================================================================================================\n\n`;

        // Loop through each email in conversation_details
        conversation.conversation_details.forEach((email, index) => {
            content += `---------------------------------------------------------------------------------------------------\n`;
            content += `Email ${index + 1}\n`;
            content += `---------------------------------------------------------------------------------------------------\n`;

            content += `From: ${email.from.name} <${email.from.address}>\n`;
            content += `To Recipients:\n`;

            email.toRecipients.forEach((recipient, i) => {
                content += `  ${i + 1}. ${recipient.name} <${recipient.address}>\n`;
            });

            if (email.ccRecipients.length > 0) {
                content += `CC Recipients:\n`;

                email.ccRecipients.forEach((recipient, i) => {
                    content += `  ${i + 1}. ${recipient.name} <${recipient.address}>\n`;
                });
            }

            content += `\nSent: ${new Date(email.sentDateTime).toLocaleString()}\n`;
            content += `Importance: ${email.importance.charAt(0).toUpperCase() + email.importance.slice(1)}\n`;
            content += `Is External: ${email.isExternal ? 'Yes' : 'No'}\n\n`;

            content += `Message:\n`;
            content += `${email.bodyPlain_text}\n\n`;
        });

        content += `===================================================================================================\n`;
        content += `End of Conversation Subject: ${conversation.subject}\n`;
        content += `===================================================================================================\n\n`;
    });

    fs.writeFileSync(filePath, content, 'utf8');
    console.log(`File saved successfully at ${filePath}`);
}

function extractTextFromHTML(htmlContent) {
    // Convert HTML to plain text
    let plainText = htmlToText.convert(htmlContent, {
        wordwrap: false,
        selectors: [
            { selector: 'img', format: 'skip' }, // Skip images
            { selector: 'a', options: { hideLinkHrefIfSameAsText: true } } // Handle links
        ]
    });

    // Split the text at the "CONFIDENTIALITY NOTICE:" keyword and keep only the part before it
    let splitText = plainText.split(/CONFIDENTIALITY NOTICE:/i);
    plainText = splitText[0].trim(); // Take the text before "CONFIDENTIALITY NOTICE:" and trim it

    // Further split the text at the "From" keyword and keep only the part before it
    splitText = plainText.split(/From:/i);
    plainText = splitText[0].trim(); // Take the text before "From" and trim it

    return plainText;
}


const generateEmailPreview = (textContent, maxLength = 150) => {
    // Trim the content to the specified max length and add ellipsis if necessary
    if (textContent.length > maxLength) {
        return textContent.substring(0, maxLength) + '...';
    }
    return textContent;
};

function getCurrentFormattedDateTime() {
    const now = new Date();

    // Options for formatting the date and time
    const options = {
        month: 'short',
        day: '2-digit',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        hour12: false
    };

    // Formatting the date and time
    const formattedDateTime = now.toLocaleDateString('en-US', options).replace(',', '');

    return formattedDateTime;
}

module.exports = {
    checkFileExists,
    convertJSONToRender,
    convertFormatJSON,
    renderMultipleObjects,
    convertJSONToTXT,
    removeEmptyObjects,
    isArrayOfObjects,
    removeSpecialCharactersWithSpace,
    hasSingleSpecialCharacter,
    splitBySquareBrackets,
    removeDuplicateEmails,
    sortUsersAlphabetically,
    removeDuplicatesBasedOnEmail,
    extractTextFromHTML,
    generateEmailPreview,
    saveEmailsToTxtFile,
    getCurrentFormattedDateTime
};
