# excel-related-table-tool

Tool for speeding up adding multiple VLOOKUPs between two tables.

## Description

The tool stores the key column in the source table, and the destination table, and saves it on the worksheet of the destination table, in a named range.

Menu items are added to the right click menu, allowing the user to replace the contents of an existing column with a VLOOKUP to data in the source column. Since we have already stored the key columns in both tables, the user only needs to pick the column from the source table.

The user can also insert new columns into their destination table. A dialog box presents them with a list of all the columns in the source table, and they can use checkboxes to create multiple new columns with the VLOOKUP already applied.

## TODO

- [x] Map source and destination table
- [x] Replace column with VLOOKUP to source
- [x] Insert multiple new columns in Destination table
- [ ] Prompt user to open source file if closed, and fail gracefully if necessary
- [ ] Prevent replacing key column in destination table
- [ ] Ensure right click menu fails gracefully in unlinked worksheets
- [ ] Allow user to insert XLOOKUP, VLOOKUP, INDEX(MATCH), =R1C1, or Values
- [ ] Allow user to choose which formatting to import
- [ ] Better handling for renaming destination columns if name already exists
- [ ] Track which columns are already linked, and do not insert duplicates
- [ ] Unit tests
- [ ] More error-handling