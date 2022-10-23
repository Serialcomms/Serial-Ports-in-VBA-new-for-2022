## Ribbon Customisation

<p float="left">
  <img align="top" src="/Ribbon/COM_PORT_TAB.png" alt="COM_PORT TAB" title="COM Port Ribbon Tab" width="30%" height="30%">
 <img src="/Ribbon/Access-Only/ACCESS_RIBBON.png" alt="Access Ribbon" title="Access Ribbon Tab" width="30%" height="30%">
</p>





#### Adding custom Ribbon tabs and commands

<details><summary>Ribbon Editor (Word and Excel)</summary>
<p>

The [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor/releases/tag/v1.9.0) is recommended for Ribbon customisation.  

Download and install RibbonX following the instructions provided with it.  

Download the file `RIBBON_2010.xml` from this folder in preparation for use.  

Follow the [instructions](How-To.md) to install the `RIBBON_2010.xml` sample customisation file.

</p>
</details>

<details><summary>Ribbon Customisation (Access only)</summary>
<p>
 
The RibbonX Editor should **not** be used to modify the Ribbon in Access documents. 
 
Instead, Microsoft instructions [here](https://support.microsoft.com/en-us/office/create-a-custom-ribbon-in-access-45e110b9-531c-46ed-ab3a-4e25bc9413de) and [here](https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/how-to-apply-a-custom-ribbon-when-starting-access)
detail how to create and apply custom Ribbons in Access. 

Local instruction summary [here](Access-Only/README_ACCESS.md) 
 
</p>
</details>
 

#### Adding Office and Custom Ribbon icons

<details><summary>Using Office Icons</summary>
<p>

A list of icons included with Office is available here [Microsoft Office Icon Gallery Download](https://www.microsoft.com/en-nz/download/confirmation.aspx?id=21103)

Further information can be found online by searching for *msoImage*

Ribbon Office icons can be changed by editing the required XML file section in RibbonX, e.g. `imageMso="NewOfficeIconName"` 
 
</p>
</details> 

<details><summary>Using Custom Icons</summary>
<p>

Custom icons can also be added from RibbonX. Use the **Insert > Icons** menu option to add a new icon file to the document. 

Ribbon Custom icons can be changed by editing the required XML file section in RibbonX, e.g. `image="MyCustomIconName"` 

The following image filetypes can be used. Image size should be between 16 x 16 to 128 x 128 

 .bmp  
 .gif   
 .jpg  
 .png  

The filetype suffix should not be included in the XML 
 
Check online for further information on supported icon types and sizes for your Office version.
 
</p>
</details> 


#### Application Development

Further VBA and Ribbon XML development is required to make the final document suitable for your intended use.  

