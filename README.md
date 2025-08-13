# CM360 Audit System

## Overview
A comprehensive Google Apps Script system for auditing CM360 (Campaign Manager 360) campaigns with advanced exclusions management and automated reporting capabilities.

## Features
- **Campaign Auditing**: Automatically audits CM360 campaigns across 9 team configurations
- **Exclusions Management**: Integrated exclusions sheet with automatic population and validation
- **Placement Name Lookup**: Updates placement names from CM360 reports with data validation
- **Excel Export**: Generates detailed audit reports with proper formatting
- **Team Configuration**: Supports multiple team setups with customizable audit parameters

## Key Components
- `Code.js` - Main audit system (1800+ lines)
- `ConfigPicker.html` - Team configuration interface
- `Dashboard.html` - Main dashboard interface
- `appsscript.json` - Google Apps Script manifest

## Setup Instructions
1. Create a new Google Apps Script project
2. Copy the contents of `Code.js` into the script editor
3. Add the HTML files for the dashboard and configuration picker
4. Configure CM360 API access and credentials
5. Set up the exclusions sheet for placement filtering

## Usage
1. **Dashboard Access**: Open the script and run the dashboard function
2. **Team Selection**: Use the Config Picker to select your team configuration
3. **Run Audit**: Execute the audit process for your selected campaigns
4. **Exclusions Management**: Use the exclusions sheet to filter specific placements
5. **Update Placement Names**: Use the "Update Placement Names" button to refresh placement data
6. **Export Results**: Download audit results as formatted Excel files

## Exclusions Sheet Features
- Automatic creation and population
- Real-time validation during edits
- Integration with audit logic
- Support for placement-specific exclusions

## Technical Details
- Built for Google Apps Script environment
- Integrates with CM360 API
- Uses Google Sheets for data management
- Supports Excel export with custom formatting
- Includes comprehensive error handling and validation

## Version History
- Enhanced with exclusions management system
- Added placement name lookup with validation
- Improved UI with emoji fixes
- Integrated automated sheet management

## Support
This system is designed for enterprise-level CM360 campaign auditing with advanced exclusions management capabilities.
