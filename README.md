# OpenWeather Data Pipeline
Idempotent API-driven weather ingestion and automated reporting system implementing configuration-driven execution, structured error handling, composite-key deduplication, and scheduler-based deployment.

Key Features
Configuration-Driven Ingestion

• Reads cities and recipients from Excel-based configuration
• Dynamically constructs API requests (city / latitude-longitude)
• Environment-based credential management

Resilient API Handling

• HTTP status code classification (401, 404, 429, 500)
• Defensive request validation
• Structured logging to prevent silent failures

Dynamic JSON Parsing

• Reusable nested JSON traversal utilities
• Handles mixed dictionary/list structures
• Avoids hardcoded extraction paths

Schema Normalization

• Extracts and structures:
○ Temperature
○ Humidity
○ Wind Speed
○ Weather Conditions
○ Coordinates
• Converts raw API payload into analytics-ready dataset

Idempotent Execution

• Composite key enforcement (Date + City) for safe re-execution
• Idempotent email dispatch using delivery log (Date + Receiver + Status)
• Prevents duplicate records and redundant daily email deliveries

Multi-Layer Storage Architecture
Layer	Purpose
Weather_Report.xlsx	Consolidated daily dataset
Log_File.xlsx	Delivery state & execution tracking
Configuration Files	Ingestion & distribution control
