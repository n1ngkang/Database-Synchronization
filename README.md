# Contract Data Synchronization Engine
## Executive Summary
This project is a Google Apps Script (GAS) synchronization engine engineered to resolve persistent data integrity and execution timeout issues encountered within large-scale, cross-spreadsheet data flows, resolveing the instability and slow performance problems brought by previous inefficient scripting attempts. 
This engine completely automates the transfer of critical Subscription Contract data from a primary Source Database to a Public View Layer, achieving robust and 24/7 automation via a two-hour time-based trigger cycle. 
## Core Technical Contributions
### Timeout Mitigation via Asynchronous Trigger Chaining
* Batch Processing: To circumvent the 6-minute GAS execution timeout limit, the primary data transfer function ```import_RSV_Plans_Final``` employs data chunking (3,000-row batches) when moving raw data to ensure stable data I/O.
* Decoupling with Trigger Chaining: After the heavy data import, the script creates an asynchronous trigger ```ScriptApp.newTrigger``` with a 10-second delay to launch the final processing function ```standardizeData```.
  This strategy decouples the intensive write operation from the final calculation phase, ensuring procedural stability.
### Performance Optimization: Formula Lock-In
* Dynamic Calculation and Restoration: Initially, after the batch transfer of raw data, the system actively restores necessary formulas (Columns AJ:AT) to allow for dependent data mapping and accurate calculation against other sheet references.
* Formula-to-Value Conversion: Once the calculations are complete, the standardizeData function immediately converts the results of all calculated fields (Columns AJ:AT) from live formulas to static numerical values using the ```setValues(values)``` method.
  This locks the final result, significantly enhancing stability and user responsiveness.
### Data Fidelity and Standardization
* Integrity Focus: I ensured the final data snapshot is reliable by using the Formula Lock-In strategy. This is crucial for verifying irregular numerical values of subscription.
* Format Enforcement: The script enforces strict formatting, thereby enhancing data fidelity and ensuring accurate representation of subscription details.
## Setup and Deployment
This project utilizes modern version control practices and is deployed via Clasp.
* Prerequisites: Node.js, npm, and the Google Apps Script Command Line Interface (Clasp).
* Installation: Install Clasp globally via npm: ```npm install -g @google/clasp```.
* Authentication: Authenticate with your Google account: ```clasp login```.
* Deployment: The project is deployed by pushing the local code to the linked GAS project: ```clasp push```.
