# Branch Master Project

This project processes data from an Excel file and inserts or updates records into a PostgreSQL database. It is designed to handle branch data for a particular organization, which includes branch details such as name, address, region, state, and other relevant information.

## Features

- Processes Excel data from an input file.
- Checks if a branch exists in the database by `branch_id`.
- If the branch exists, it updates the record if there are changes.
- If the branch does not exist, it inserts a new record.
- Handles auto-increment of the serial number (`s_no`) for new records.
- Stores logs for processed entries and errors.

## Prerequisites

Before running the project, ensure you have the following installed:

- **Node.js** (v14.x or higher)
- **PostgreSQL** (v12.x or higher)
- **Excel file** formatted with branch data

- run the branch_masterV2.js

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/your-username/branch-master.git
