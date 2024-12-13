const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");
const { Pool } = require("pg");

// Load environment variables from a .env file (optional, for local development)
require('dotenv').config();

// PostgreSQL connection setup

const pool = new Pool({
    user: process.env.DB_USER,
    host: process.env.DB_HOST,
    database: process.env.DB_NAME,
    password: process.env.DB_PASSWORD,
    port: process.env.DB_PORT || 5432, // Default to 5432 if not set
  });

  // Helper function to split value into ID and Name
function splitValue(value) {
    const parts = value.split('-');
    if (parts.length === 2) {
      return { id: parts[0], name: parts[1] };
    } else {
      return { id: '', name: value }; // If no hyphen, use the full value as the name
    }
  }
// Function to process Excel data
async function processExcelFile(filePath) {
    // const client = await pool.connect();
  try {
    // await client.query('BEGIN'); // Start transaction
    // Read Excel file
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = xlsx.utils.sheet_to_json(sheet);

    for (const row of rows) {
      const {
        State,
        Region,
        Area,
        Branch,
      } = row;
    // Extract branch_id (assuming it's the part before the first hyphen)
    // const branch_id = Branch.split('-')[0]; // This will give "B110" from "B110-Darbhanga"
    // const branch_name = Branch.split('-')[1];
    // const area_id = Area.split('-')[0];
    // const area_name = Area.split('-')[1];
    // const region_id = Region.split('-')[0];
    // const region_name = Region.split('-')[1];
    // Handle Branch
    const { id: branch_id, name: branch_name } = splitValue(Branch);

    // Handle Area
    const { id: area_id, name: area_name } = splitValue(Area);

    // Handle Region
    const { id: region_id, name: region_name } = splitValue(Region);

    
      // Check if branch_id exists in the database
      const { rows: existingRows } = await pool.query(
        "SELECT * FROM public.sample_branch WHERE branch_id = $1",
        [branch_id]
      );

      if (existingRows.length > 0) {
        // Branch ID exists, check for differences
        const dbRow = existingRows[0];
     
        const isSame =
          dbRow.area_id === area_id &&
          dbRow.area_name === area_name &&
          dbRow.region_id === region_id &&
          dbRow.region_name === region_name &&
          dbRow.state === State;

        if (!isSame) {
          // Update database row if data is different
          await pool.query(
            `UPDATE public.sample_branch
             SET branch_name = $1, area_id = $2, area_name = $3, region_id = $4, region_name = $5 , state = $6
             WHERE branch_id = $7`,
            [branch_name, area_id, area_name, region_id, region_name,State, branch_id]
          );
          console.log(`Updated branch_id: ${branch_id}`);
        } else {
          console.log(`No changes for branch_id: ${branch_id}`);
        }
      } else {
        // Insert new row if branch_id does not exist
        await pool.query(
          `INSERT INTO public.sample_branch (branch_id, branch_name, area_id, area_name, region_id, region_name,state)
           VALUES ($1, $2, $3, $4, $5, $6,$7)`,
          [branch_id, branch_name, area_id, area_name, region_id, region_name,State]
        );
        console.log(`Inserted new branch_id: ${branch_id}`);
      }
    //   await client.query('COMMIT'); // Commit transaction
    }
  } catch (error) {
    console.error("Error processing Excel file:", error);
  } finally {
    // Close the database connection
    await pool.end();
  }
}

// Path to your Excel file
const filePath = path.join(process.env.file_path);

// Process the Excel file
processExcelFile(filePath);
