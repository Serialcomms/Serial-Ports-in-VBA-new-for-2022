# Ribbon Customisation in Access


Steps to customise the Ribbon in Access only are summarised below

1. Open new or existing Access document
2. Check if a database table **`USysRibbons`** exists first
3. If not, then create a new table **`USysRibbons`** 
4. New Table Columns = `ID (AutoNumber) , RibbonName (Short Text) , RibbonXml (Long Text)`
5. Insert a new row into table `USysRibbons , RibbonName = "COM_1"`
6. Update the new row by copying contents of file `RIBBON_ACCESS.xml` into column `RibbonXml`
7. Save, Close and Re-Open the Access document
8. Navigate to File > Options > Current Database > Ribbon and Toolbar Options > Ribbon Name
9. Select **COM_1** and click OK
10. Save, Close and Re-Open the Access document again
11. Check that a new Ribbon tab COM Port 1 is now present
12. Import modules SERIAL_PORT_VBA and SERIAL_PORT_RIBBON if not done previously
13. Test to ensure that Ribbon icons work and COM Port operation is successful
