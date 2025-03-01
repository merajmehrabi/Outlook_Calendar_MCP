/**
 * test-outlook-connection.js - Comprehensive test script for Outlook calendar operations
 * 
 * This script tests if the MCP tool can access and manipulate your Outlook calendar by:
 * 1. Testing connection by listing available calendars
 * 2. Testing READ operation by listing events
 * 3. Testing WRITE operation by creating a test event
 * 4. Testing DELETE operation by removing the test event
 */

import { exec } from 'child_process';
import path from 'path';
import { fileURLToPath } from 'url';
import { promisify } from 'util';

// Convert exec to promise-based
const execPromise = promisify(exec);

// Get the directory name of the current module
const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Script paths
const scriptsDir = path.join(__dirname, 'scripts');
const getCalendarsScript = path.join(scriptsDir, 'getCalendars.vbs');
const listEventsScript = path.join(scriptsDir, 'listEvents.vbs');
const createEventScript = path.join(scriptsDir, 'createEvent.vbs');
const updateEventScript = path.join(scriptsDir, 'updateEvent.vbs');
const deleteEventScript = path.join(scriptsDir, 'deleteEvent.vbs');

// Test event details
const testEventSubject = `Test Event ${new Date().toISOString()}`;
const today = new Date();
const formattedDate = `${today.getMonth() + 1}/${today.getDate()}/${today.getFullYear()}`;
let testEventId = null;

// Helper function to parse script output
function parseScriptOutput(stdout) {
  if (stdout.includes('SUCCESS:')) {
    const jsonStr = stdout.substring(stdout.indexOf('SUCCESS:') + 'SUCCESS:'.length).trim();
    try {
      return { success: true, data: JSON.parse(jsonStr) };
    } catch (parseError) {
      return { success: false, error: `Error parsing JSON: ${parseError.message}`, raw: jsonStr };
    }
  } else if (stdout.includes('ERROR:')) {
    const errorMessage = stdout.substring(stdout.indexOf('ERROR:') + 'ERROR:'.length).trim();
    return { success: false, error: errorMessage };
  } else {
    return { success: false, error: 'Unexpected script output', raw: stdout };
  }
}

// Execute a VBScript with parameters
async function executeScript(scriptPath, params = []) {
  const paramString = params.map(p => `/${p.name}:"${p.value}"`).join(' ');
  const command = `cscript //NoLogo "${scriptPath}" ${paramString}`;
  
  try {
    const { stdout, stderr } = await execPromise(command);
    
    if (stderr) {
      return { success: false, error: stderr };
    }
    
    return parseScriptOutput(stdout);
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// Main test function
async function runTests() {
  console.log('🔍 TESTING OUTLOOK CALENDAR MCP TOOL');
  console.log('====================================');
  
  // Test 1: Connection Test
  console.log('\n📋 TEST 1: Connection Test (getCalendars.vbs)');
  console.log('-------------------------------------------');
  
  const connectionResult = await executeScript(getCalendarsScript);
  
  if (connectionResult.success) {
    console.log('✅ Connection test passed!');
    console.log('Available calendars:');
    console.log(JSON.stringify(connectionResult.data, null, 2));
  } else {
    console.error('❌ Connection test failed:', connectionResult.error);
    console.error('Cannot proceed with further tests without Outlook connection.');
    process.exit(1);
  }
  
  // Test 2: Read Test
  console.log('\n📋 TEST 2: Read Test (listEvents.vbs)');
  console.log('-------------------------------------------');
  
  const readResult = await executeScript(listEventsScript, [
    { name: 'startDate', value: formattedDate },
    { name: 'endDate', value: formattedDate }
  ]);
  
  if (readResult.success) {
    console.log('✅ Read test passed!');
    console.log(`Found ${readResult.data.length} events for today (${formattedDate}).`);
  } else {
    console.error('❌ Read test failed:', readResult.error);
    console.error('Cannot proceed with further tests.');
    process.exit(1);
  }
  
  // Test 3: Write Test
  console.log('\n📋 TEST 3: Write Test (createEvent.vbs)');
  console.log('-------------------------------------------');
  
  const writeResult = await executeScript(createEventScript, [
    { name: 'subject', value: testEventSubject },
    { name: 'startDate', value: formattedDate },
    { name: 'startTime', value: '2:00 PM' },
    { name: 'endDate', value: formattedDate },
    { name: 'endTime', value: '2:30 PM' },
    { name: 'location', value: 'Test Location' },
    { name: 'body', value: 'This is a test event created by the Outlook Calendar MCP Tool test script.' }
  ]);
  
  if (writeResult.success) {
    testEventId = writeResult.data.eventId;
    console.log('✅ Write test passed!');
    console.log(`Created test event with ID: ${testEventId}`);
  } else {
    console.error('❌ Write test failed:', writeResult.error);
    console.error('Cannot proceed with delete test.');
    process.exit(1);
  }
  
  // Test 4: Update Test
  console.log('\n📋 TEST 4: Update Test (updateEvent.vbs)');
  console.log('-------------------------------------------');
  
  if (!testEventId) {
    console.error('❌ Update test skipped: No event ID from write test.');
    process.exit(1);
  }
  
  const updateResult = await executeScript(updateEventScript, [
    { name: 'eventId', value: testEventId },
    { name: 'subject', value: `${testEventSubject} - UPDATED` },
    { name: 'location', value: 'Updated Test Location' }
  ]);
  
  if (updateResult.success && updateResult.data.success) {
    console.log('✅ Update test passed!');
    console.log(`Successfully updated test event with ID: ${testEventId}`);
  } else {
    console.error('❌ Update test failed:', updateResult.error || 'Unknown error');
    process.exit(1);
  }
  
  // Test 5: Delete Test
  console.log('\n📋 TEST 5: Delete Test (deleteEvent.vbs)');
  console.log('-------------------------------------------');
  
  if (!testEventId) {
    console.error('❌ Delete test skipped: No event ID from write test.');
    process.exit(1);
  }
  
  const deleteResult = await executeScript(deleteEventScript, [
    { name: 'eventId', value: testEventId }
  ]);
  
  if (deleteResult.success && deleteResult.data.success) {
    console.log('✅ Delete test passed!');
    console.log(`Successfully deleted test event with ID: ${testEventId}`);
  } else {
    console.error('❌ Delete test failed:', deleteResult.error || 'Unknown error');
    process.exit(1);
  }
  
  // Summary
  console.log('\n📋 TEST SUMMARY');
  console.log('-------------------------------------------');
  console.log('✅ Connection Test: PASSED');
  console.log('✅ Read Test: PASSED');
  console.log('✅ Write Test: PASSED');
  console.log('✅ Update Test: PASSED');
  console.log('✅ Delete Test: PASSED');
  console.log('\n🎉 All tests passed! The Outlook Calendar MCP Tool is working correctly.');
}

// Run the tests
runTests().catch(error => {
  console.error('Error running tests:', error);
  process.exit(1);
});
