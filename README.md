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
