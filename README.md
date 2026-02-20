# OpenWeather Data Pipeline
Idempotent API-driven weather ingestion and automated reporting system implementing configuration-driven execution, structured error handling, composite-key deduplication, and scheduler-based deployment.

Key Features
1. Configuration-Driven Ingestion
• Reads cities and recipients from Excel-based configuration
• Dynamically constructs API requests using city names
• Environment-based credential management(api_key and app_password)

2. Resilient API Handling
• HTTP status code classification (401, 404, 429, 500...)
• Defensive request validation
• Structured logging to prevent silent failures

3. Dynamic JSON Parsing
• Reusable nested JSON traversal utilities
• Avoids hardcoded extraction paths

4. Schema Normalization
• Extracts and structures:
○ Temperature
○ Humidity
○ Wind Speed
○ Weather Conditions
○ Coordinates
• Converts raw API payload into analytics-ready dataset

5. Idempotent Execution
• Idempotent weather report using Composite keys (Country + City + Date)
• Idempotent email dispatch using delivery log (Receiver + Status + Date)
• Prevents duplicate records and redundant daily email deliveries

6. Multi-Layer Storage Architecture
Weather_Report.xlsx	Consolidated daily dataset
Log_File.xlsx	Delivery state & execution tracking
Configuration Files: Receiver Mail + City Names List
