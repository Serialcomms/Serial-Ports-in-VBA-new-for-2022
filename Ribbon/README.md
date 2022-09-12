## Ribbon Customisation

#### Adding custom Ribbon tabs and commands

<details><summary>Ribbon Editor</summary>
<p>

The [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor/releases/tag/v1.9.0) is recommended for Ribbon customisation.  

Download and install RibbonX following the instructions provided with it.  

Download the file RIBBON_2010.xml from this folder in preparation for use.  

Follow the [HowTo](How-To.md) instructions to install the RIBBON_2010.xml sample customisation file.

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


#### VBA Development

Further VBA development is required to make the final document suitable for your application.  

