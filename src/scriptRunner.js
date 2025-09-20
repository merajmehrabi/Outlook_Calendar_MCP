/**
 * scriptRunner.js - Handles execution of VBScript files and processes their output
 */

import { exec } from 'child_process';
import path from 'path';
import { fileURLToPath } from 'url';

// Get the directory name of the current module
const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Constants
const SCRIPTS_DIR = path.resolve(__dirname, '../scripts');
const SUCCESS_PREFIX = 'SUCCESS:';
const ERROR_PREFIX = 'ERROR:';

/**
 * Executes a VBScript file with the given parameters
 * @param {string} scriptName - Name of the script file (without .vbs extension)
 * @param {Object} params - Parameters to pass to the script
 * @returns {Promise<Object>} - Promise that resolves with the script output
 */
export async function executeScript(scriptName, params = {}) {
  return new Promise((resolve, reject) => {
    // Build the command with UTF-8 support
    const scriptPath = path.join(SCRIPTS_DIR, `${scriptName}.vbs`);
    let command = `chcp 65001 >nul 2>&1 && cscript //NoLogo "${scriptPath}"`;
    
    // Add parameters
    for (const [key, value] of Object.entries(params)) {
      if (value !== undefined && value !== null && value !== '') {
        // Handle special characters in values
        const escapedValue = value.toString().replace(/"/g, '\\"');
        command += ` /${key}:"${escapedValue}"`;
      }
    }
    
    // Execute the command with UTF-8 encoding
    exec(command, { encoding: 'utf8' }, (error, stdout, stderr) => {
      // Check for execution errors
      if (error && !stdout.includes(SUCCESS_PREFIX)) {
        return reject(new Error(`Script execution failed: ${error.message}`));
      }
      
      // Check for script errors
      if (stdout.includes(ERROR_PREFIX)) {
        const errorMessage = stdout.substring(stdout.indexOf(ERROR_PREFIX) + ERROR_PREFIX.length).trim();
        return reject(new Error(`Script error: ${errorMessage}`));
      }
      
      // Process successful output
      if (stdout.includes(SUCCESS_PREFIX)) {
        try {
          const jsonStr = stdout.substring(stdout.indexOf(SUCCESS_PREFIX) + SUCCESS_PREFIX.length).trim();
          const result = JSON.parse(jsonStr);
          return resolve(result);
        } catch (parseError) {
          return reject(new Error(`Failed to parse script output: ${parseError.message}`));
        }
      }
      
      // If we get here, something unexpected happened
      reject(new Error(`Unexpected script output: ${stdout}`));
    });
  });
}

/**
 * Lists calendar events within a specified date range
 * @param {string} startDate - Start date in MM/DD/YYYY format
 * @param {string} endDate - End date in MM/DD/YYYY format (optional)
 * @param {string} calendar - Calendar name (optional)
 * @returns {Promise<Array>} - Promise that resolves with an array of events
 */
export async function listEvents(startDate, endDate, calendar) {
  return executeScript('listEvents', { startDate, endDate, calendar });
}

/**
 * Creates a new calendar event
 * @param {Object} eventDetails - Event details
 * @returns {Promise<Object>} - Promise that resolves with the created event ID
 */
export async function createEvent(eventDetails) {
  return executeScript('createEvent', eventDetails);
}

/**
 * Finds free time slots in the calendar
 * @param {string} startDate - Start date in MM/DD/YYYY format
 * @param {string} endDate - End date in MM/DD/YYYY format (optional)
 * @param {number} duration - Duration in minutes (optional)
 * @param {number} workDayStart - Work day start hour (0-23) (optional)
 * @param {number} workDayEnd - Work day end hour (0-23) (optional)
 * @param {string} calendar - Calendar name (optional)
 * @returns {Promise<Array>} - Promise that resolves with an array of free time slots
 */
export async function findFreeSlots(startDate, endDate, duration, workDayStart, workDayEnd, calendar) {
  return executeScript('findFreeSlots', {
    startDate,
    endDate,
    duration,
    workDayStart,
    workDayEnd,
    calendar
  });
}

/**
 * Gets the response status of meeting attendees
 * @param {string} eventId - Event ID
 * @param {string} calendar - Calendar name (optional)
 * @returns {Promise<Object>} - Promise that resolves with meeting details and attendee status
 */
export async function getAttendeeStatus(eventId, calendar) {
  return executeScript('getAttendeeStatus', { eventId, calendar });
}

/**
 * Deletes a calendar event by its ID
 * @param {string} eventId - Event ID
 * @param {string} calendar - Calendar name (optional)
 * @returns {Promise<Object>} - Promise that resolves with the deletion result
 */
export async function deleteEvent(eventId, calendar) {
  return executeScript('deleteEvent', { eventId, calendar });
}

/**
 * Updates an existing calendar event
 * @param {string} eventId - Event ID to update
 * @param {string} subject - New subject (optional)
 * @param {string} startDate - New start date in MM/DD/YYYY format (optional)
 * @param {string} startTime - New start time in HH:MM AM/PM format (optional)
 * @param {string} endDate - New end date in MM/DD/YYYY format (optional)
 * @param {string} endTime - New end time in HH:MM AM/PM format (optional)
 * @param {string} location - New location (optional)
 * @param {string} body - New body/description (optional)
 * @param {string} calendar - Calendar name (optional)
 * @returns {Promise<Object>} - Promise that resolves with the update result
 */
export async function updateEvent(eventId, subject, startDate, startTime, endDate, endTime, location, body, calendar) {
  return executeScript('updateEvent', {
    eventId,
    subject,
    startDate,
    startTime,
    endDate,
    endTime,
    location,
    body,
    calendar
  });
}

/**
 * Lists available calendars
 * @returns {Promise<Array>} - Promise that resolves with an array of calendars
 */
export async function getCalendars() {
  return executeScript('getCalendars');
}
