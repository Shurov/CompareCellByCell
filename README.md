# CompareCellByCell

This VBA project helps us to automate some routine tasks at work - compare 2 Excel worksheets or CSV-based text files cell-by-cell to find deviations in an exact line and an exact field.

Project structure:
- *'Forms'* folder - forms to be imported into the project. *.frm* files are treated as text files to track differences between commits
- *'Modules'* folder - modules to be imported. *.bas* files are treated as text files as well.
- *'Resourses'* folder:
   - *V&T Ribbon.xlam* - VBA-code as an Excel Addin. Ready for import into the system.
   - *Setup V&T ribbon.xls* - performs automated installation of the Excel Addin into the system. Need Addin to be present in the same dir.
   - *macro_code.xlsm* - VBA-code as a macro-workbook. Convenient to change the code in such form. Needs to be saved as an Excel Addin to work.
- *customUI14.xml* - XML markup that creates Excel-ribbon with buttons to access the code. Custom UI Editor (http://openxmldeveloper.org/archive/2006/05/25/CustomUIeditor.aspx) required to modify *.xlsm" files.
