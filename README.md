## DET Data Tools

Current tools:

1. GN-RP_AccessToExcel converts Access Database for a specific project into an xlsx file with the same data.

2. GN-RP_ConvertDETToFlatWorkbook processes the indicated DET into a flat (data only, formatting cleared) workbook of the project name.
  Note the *.xlsm file name is assumed to be (GN/RP)" "(Project)" "(Date)" "(pre/pub).xlsm

3. GN-RP_EnteredDataVerification.xlsm checks that the indicated DET has data in the proper format. This tool checks for spaces, missing data, and so on.

4. GN-RP_FlatWorkbookToCombinedFlatWorkbook takes the indicated previously flattened DET workbooks and combines them into a single flattened workbook.
  Checks for minor errors, missing data, and so on.
  
Note: these tools are written in Excel VBA.
New tools will eventually be available to perform other desired functionality.