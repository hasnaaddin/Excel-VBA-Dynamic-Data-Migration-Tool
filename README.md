# Excel VBA Dynamic Data Migration & Reconciliation Tool

A configuration-driven Excel VBA solution for migrating data from source datasets into structured target templates.

The tool uses dynamic source-to-target column mappings, configurable data types and formats, automated default values, reconciliation checks, and whitespace validation to provide a controlled and repeatable data migration process.

## Key Features

- Dynamic source-to-target column mapping
- Single source worksheet to multiple target worksheet migration
- Configuration-driven migration with no VBA changes required for new mappings
- Supports TEXT, NUMERIC and DATE data types
- Configurable date and numeric formatting
- Automatic population of target-only fields using default values
- Source vs target record-count reconciliation
- Numeric source vs target sum reconciliation
- Automatic Match / No Match validation
- Detection and cleansing of leading, trailing and duplicate spaces
- Missing source and target column reporting
- Timestamped migrated output files
- Direct link to the latest generated migration file
- Modular VBA architecture for easier maintenance

## How It Works

The migration process is controlled through the Excel `Control` worksheet rather than hard-coded source-to-target mappings.

1. Configure the source and target worksheet structure in the Control sheet.
2. Define the source-to-target column mappings.
3. Specify data types, formats and optional default values.
4. Click **Run Migration** from the Frontpage.
5. Select the source dataset and target template.
6. VBA dynamically reads the configuration and migrates the required data.
7. A single source worksheet can populate multiple worksheets within the target workbook.
8. The tool performs automated count and numeric sum reconciliation.
9. Whitespace issues are detected, reported and cleansed.
10. A timestamped migrated workbook is generated and linked from the Frontpage.

The original target workbook acts as the import template, while the populated migration output is saved as a new timestamped file.
