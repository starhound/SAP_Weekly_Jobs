# SAP_Weekly_Jobs
Generates a Excel report from SAP B1 data

Requires the addition of a "Job_Schedule_ReportPhase" stored procedure on your MS-SQL or HANA server (note - HANA untested).

Shows all jobs for the time range desired that have been completed but yet to be invoiced or batched.

Requires a [SAP B1](https://www.sap.com/products/business-one) deployment with the [Eralis Jobs](https://eralis.software/products/job-suite/) addon. 

I don't believe it requires any special UDF's, but will require some re-factoring to have this program producing desired results for your server.

Requires a connection string be placed on SAP_DATA_PULL.cs on line 39.
