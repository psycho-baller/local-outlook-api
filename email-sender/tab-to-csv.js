/**
 * Tab-to-CSV Converter
 * 
 * This script converts a tab-delimited file to a semicolon-separated file
 * with double quotes around field values.
 * 
 * Usage: node tab-to-csv.js <input-file> <output-file>
 */

const fs = require('fs');
const path = require('path');
const readline = require('readline');

// Get command line arguments
const args = process.argv.slice(2);
const inputFile = args[0];
const outputFile = args[1] || inputFile.replace(/\.[^.]+$/, '') + '.csv';

// Validate arguments
if (!inputFile) {
  console.error('Error: Input file is required');
  console.log('Usage: node tab-to-csv.js <input-file> <output-file>');
  process.exit(1);
}

/**
 * Convert a tab-delimited line to a semicolon-separated line with quoted values
 * @param {string} line - Tab-delimited line
 * @returns {string} - Semicolon-separated line with quoted values
 */
function convertLine(line) {
  // Split the line by tabs
  const fields = line.split('\t');
  
  // Process each field: trim whitespace and wrap in double quotes
  const processedFields = fields.map(field => {
    // Trim the field
    const trimmedField = field.trim();
    
    // Escape any double quotes in the field by doubling them
    const escapedField = trimmedField.replace(/"/g, '""');
    
    // Wrap the field in double quotes
    return `"${escapedField}"`;
  });
  
  // Join the fields with semicolons
  return processedFields.join(';');
}

async function convertFile() {
  try {
    // Check if input file exists
    if (!fs.existsSync(inputFile)) {
      console.error(`Error: Input file '${inputFile}' does not exist`);
      process.exit(1);
    }
    
    console.log(`Converting '${inputFile}' to '${outputFile}'...`);
    
    // Create read stream and interface
    const fileStream = fs.createReadStream(inputFile);
    const rl = readline.createInterface({
      input: fileStream,
      crlfDelay: Infinity
    });
    
    // Create write stream
    const writeStream = fs.createWriteStream(outputFile);
    
    // Process each line
    let lineCount = 0;
    for await (const line of rl) {
      // Convert the line
      const convertedLine = convertLine(line);
      
      // Write the converted line to the output file
      writeStream.write(convertedLine + '\n');
      
      lineCount++;
    }
    
    // Close the write stream
    writeStream.end();
    
    console.log(`Conversion complete! Processed ${lineCount} lines.`);
    console.log(`Output saved to: ${path.resolve(outputFile)}`);
  } catch (error) {
    console.error('Error converting file:', error.message);
    process.exit(1);
  }
}

// Run the conversion
convertFile();
