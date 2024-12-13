/*
get branch_id from excel
if(exists in db){
check entrys for updation orelse ignore
}
else{
insert as new entry
}

****
To fix:: no auto_increment of id upon new entry into db
*/

const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");
const { Pool } = require("pg");

// Load environment variables from a .env file
require('dotenv').config();

// PostgreSQL connection setup
const pool = new Pool({
    user: process.env.DB_USER,
    host: process.env.DB_HOST,
    database: process.env.DB_NAME,
    password: process.env.DB_PASSWORD,
    port: process.env.DB_PORT || 5432, // Default to 5432 if not set
  });

// Get the table name from the environment variable
const tableName = process.env.DB_TABLE_NAME;

// Helper function to log messages to a file
function logToFile(message) {
  const logFilePath = path.join(__dirname, 'logs.txt'); // Path to the log file
  const timestamp = new Date().toISOString();
  fs.appendFileSync(logFilePath, `[${timestamp}] ${message}\n`, 'utf8');
}

// Function to process Excel data
async function processExcelFile(filePath) {
  // const client = await pool.connect();
  const operationLogs = []; // Array to store operation logs
  try {
    // await client.query('BEGIN'); // Start transaction
    // Read Excel file
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = xlsx.utils.sheet_to_json(sheet);
    // logToFile(rows[0]);
  
    for (const row of rows) {
      
        // logToFile(row["Branch ID"], row["State*"]);
        const {
            "S.No": serial_no,
            "Partner Bank ID": partner_bank_id,
            "Hierarchy Type*": hierarchy_type,
            "Branch ID": branch_id,
            "Branch Name*": branch_name,
            "Area ID*": area_id,
            "Area name": area_name,
            "Region ID*": region_id,
            "Region Name": region_name,
            "Cluster ID*": cluster_id,
            "Cluster Name": cluster_name,
            "Zone ID": zone_id,
            "Zone Name": zone_name,
            "State*": state_1,
            "Branch ID with name": branch_id_with_name,
            "Address 1*": address_1,
            "Address 2": address_2,
            "Pincode*": pincode,
            "State*": state_2, 
            "District*": district,
            "Email Id": email_id,
            "Phone No*": mobile_no,
            "Language": language,
            "Office Area (sq2)": office_area,
            "Date of LOI": date_of_loi,
            "Opening Date*": open_date,
            "Rent Agreement Start Date": rent_start_date,
            "Rent Effective date": rent_effective_date,
            "Rent Agreement End Date": rent_end_date,
            "Rent Per month": rent_per_month,
            "Security Deposit": security_deposit,
            "Rent Incr Percent": rent_increase_percent,
            "Branch Status": branch_status,
            "Closing Date": close_date,
            "Latitudes": latitude,
            "Longitudes": longitude,
            "Area Manager Name with Employee ID": area_manager_name,
            "Area Manager Email ID": area_manager_email,
            "Area Manager Mobile Number": area_manager_mobile,
            "Regional Manager Name with Employee ID": regional_manager_name,
            "Regional Manager Email ID": regional_manager_email,
            "Regional Manager Mobile Number": regional_manager_mobile,
            "State Head": state_head,
            "State Head Email ID": state_head_email,
            "State Head Mobile Number": state_head_mobile,
            "BC Branch Code": bc_branch_code,
            "Bank BC Branch Name": bank_bc_branch_name,
            "BC System Branch Name": bc_system_branch_name,
            "BC System Branch Code": bc_system_branch_code,
            "Admin Team Remark": admin_team_remark,
            "Virtual/Physical": virtual_physical,
            "District HQ Lat": district_hq_lat,
            "DIst HQ Lon": district_hq_lon,
            "Geopy_distance_Km": geopy_distance_km,
            "ZingHR LAT": zinghr_lat,
            "ZingHR Long": zinghr_lon,
            "FinLib Lat": finlib_lat,
            "FinLib Long": finlib_lon,
            "Spotways Lat": spotways_lat,
            "Spotways Long": spotways_lon,
          } = row;

    // Log the parsed data for debugging
    const jsOpenDate = new Date((open_date - 25569) * 86400 * 1000);
    // Set the time to 00:00:00 to remove the time part
    jsOpenDate.setHours(0, 0, 0, 0);
    // Convert to desired date format (e.g., MM/DD/YYYY)
    const formattedOpenDate = jsOpenDate.toLocaleDateString('en-GB'); 

    const jsCloseDate = new Date((close_date - 25569) * 86400 * 1000);
    jsCloseDate.setHours(0, 0, 0, 0);
    const formattedCloseDate = jsCloseDate.toLocaleDateString('en-GB'); 

    logToFile(`Processed Branch ID: ${branch_id}, State: ${state_1}`);
    logToFile(JSON.stringify({
      partner_bank_id,
      hierarchy_type,
      branch_id,
      branch_name,
      area_id,
      area_name,
      region_id,
      region_name,
      cluster_id,
      cluster_name,
      zone_id,
      zone_name,
      state_1,
      branch_id_with_name,
      address_1,
      address_2,
      pincode,
      district,
      email_id,
      mobile_no,
      language,
      formattedOpenDate,
      formattedCloseDate,
      latitude,
      longitude,
    }));

    if (!branch_id) {
        operationLogs.push({ branch_id: "N/A", operation: "Skipped (No Branch ID)" });
        continue;
    }

      if(branch_id == undefined){
        continue;
      }
      // Check if branch_id exists in the database
        const { rows: existingRows } = await pool.query(
          `SELECT * FROM ${tableName} WHERE branch_id = $1`,
          [branch_id]
        );
        
        if (existingRows.length > 0) {

        // Branch ID exists, check for differences
        const dbRow = existingRows[0];
        const isSame =
        dbRow['area_id'] === area_id &&
        dbRow['area_name'] === area_name &&
        dbRow['region_id'] === region_id &&
        dbRow['region_name'] === region_name &&
        dbRow['state'] === state_1;
      

        if (!isSame) {

          // Update database row if data is different
          await pool.query(
            `UPDATE ${tableName}
             SET branch_name = $1, 
                 area_id = $2, 
                 area_name = $3, 
                 region_id = $4, 
                 region_name = $5, 
                 state = $6,
                 partner_bank_id = $7,
                 hierarchy_type = $8,
                 cluster_id = $9,
                 cluster_name = $10,
                 zone_id = $11,
                 zone_name = $12,
                 branch_id_with_name = $13,
                 address_1 = $14,
                 address_2 = $15,
                 pincode = $16,
                 district = $17,
                 email_id = $18,
                 mobile_no = $19,
                 language = $20,
                 open_date = $21,
                 close_date = $22,
                 latitude = $23,
                 longitude = $24
             WHERE branch_id = $25`,
            [
                branch_name,
                area_id,
                area_name,
                region_id,
                region_name,
                state_1,
                partner_bank_id,
                hierarchy_type,
                cluster_id,
                cluster_name,
                zone_id,
                zone_name,
                branch_id_with_name,
                address_1,
                address_2,
                pincode,
                district,
                email_id,
                mobile_no,
                language,
                formattedOpenDate,
                formattedCloseDate,
                latitude,
                longitude,
                branch_id,
            ]
        );
        
          logToFile(`Updated branch_id: ${branch_id}`);
          operationLogs.push({ branch_id, operation: "Updated" });
        } else {
          logToFile(`No changes for branch_id: ${branch_id}`);
          operationLogs.push({ branch_id, operation: "No Changes" });

        }
      } else {
        // Get the current maximum serial number (s_no)
        const { rows } = await pool.query(`SELECT MAX(s_no) AS max_s_no FROM ${tableName}`);

        // Increment the serial number
        const maxSNo = rows[0].max_s_no || 0; // If no records, start with 0
        const newSNo = maxSNo + 1;
        // Insert new row if branch_id does not exist
        await pool.query(
          `INSERT INTO ${tableName} (
              s_no,
              branch_id, 
              branch_name, 
              area_id, 
              area_name, 
              region_id, 
              region_name,
              state, 
              partner_bank_id, 
              hierarchy_type, 
              cluster_id, 
              cluster_name, 
              zone_id, 
              zone_name, 
              branch_id_with_name, 
              address_1, 
              address_2, 
              pincode, 
              district, 
              email_id, 
              mobile_no, 
              language, 
              open_date, 
              close_date, 
              latitude, 
              longitude
          ) VALUES (
              $1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17, $18, $19, $20, $21, $22, $23, $24, $25,$26
          )`,
          [
              newSNo,
              branch_id,
              branch_name,
              area_id,
              area_name,
              region_id,
              region_name,
              state_1,
              partner_bank_id,
              hierarchy_type,
              cluster_id,
              cluster_name,
              zone_id,
              zone_name,
              branch_id_with_name,
              address_1,
              address_2,
              pincode,
              district,
              email_id,
              mobile_no,
              language,
              formattedOpenDate,
              formattedCloseDate,
              latitude,
              longitude
          ]
      );
      operationLogs.push({ branch_id, operation: "Inserted" });
        logToFile(`Inserted new branch_id: ${branch_id}`);
      }
    // await client.query('COMMIT'); // Commit transaction
    }
    // Write operation logs to an Excel file
    // writeLogsToExcel(operationLogs, "operation_logs.xlsx");
    console.log("Processing complete. Logs written to operation_logs.xlsx.");
  } catch (error) {
    console.error("Error processing Excel file:", error);
  } finally {
    // Close the database connection
    await pool.end();
  }
}


// Function to write operation logs to an Excel file
function writeLogsToExcel(logs, outputFilePath) {
  // Convert logs array to worksheet
  const worksheet = xlsx.utils.json_to_sheet(logs);

  // Create a new workbook and append the worksheet
  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, worksheet, "Logs");

  // Write the workbook to the output file
  xlsx.writeFile(workbook, outputFilePath);
}

// Path to your Excel file
const filePath = path.join(process.env.file_path);

// Process the Excel file
processExcelFile(filePath);
