# Ribbon Customisation

## Using the RibbonX Editor

_Steps below assume that the Office document to be customised has previously had the SERIAL_PORT_VBA module installed and tested._

1. Close all Office documents before continuing.
2. Open the required Office document in the RibbonX Editor.
3. Confirm that the document name appears on the left hand side.
4. From the RibbonX menu, select **Insert > Office 2010+ Office UI Part**
5. Confirm that **customUI14.xml** appears under the document name on the left.
6. Copy-and-paste contents of file **RIBBON_2010.xml** into the empty area on the right.
7. Click on **Validate** from the RibbonX editor.
8. Confirm the **Custom XML UI is well formed** message box.
9. Click **Save** from the RibbonX editor menu
10. **Close** the RibbonX editor
11. Re-open the Office document as normal in Excel/Access/Word
12. Confirm that a new tab **COM 1** is present in the document ribbon menu.
13. Confirm that the new tab contains icons similar to the image in this folder.
14. Select tab **COM Port 1** and click the **Start** icon
15. Confirm that error message 'Cannot run the macro COM_PORT_CONTROL_1' appears
16. Download the file **SERIAL_PORT_RIBBON.bas** from this folder
17. Enter the VBA Environment (Alt-F11)
18. From VBA Environment, view the Project Explorer (Control-R)
19. From Project Explorer, right-hand click and select Import File.
20. Import the file SERIAL_PORT_RIBBON.bas
21. Check that a new module **SERIAL_PORT_RIBBON** is created and visible in the Modules folder. 
22. Close and return to Office application (Alt-Q)
23. IMPORTANT - save document as type Macro-Enabled with a file name of your choice.
24. IMPORTANT - test here **assumes COM Port number 1 is available on the PC** 
25. Change line **`Const Number As Long = 1`** at start of SERIAL_PORT_RIBBON to another port number if required 
26. Re-select tab **COM Port 1** and click the **Start** icon
27. Confirm that message **Start Result=True** is displayed. 
28. Test other icons with a second device attached to the COM Port. 
