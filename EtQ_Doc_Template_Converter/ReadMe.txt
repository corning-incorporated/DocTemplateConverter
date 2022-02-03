' **************************************************
'   Copyright 2022 Corning Incorporated
'   Author: Tom Gee
' **************************************************

This solution requires Microsoft Office to be installed in order to add the following references.

The following references will need to be manually added to the solution before the project will successfully start/compile.

1. Once Microsoft Office has been installed, right-click References from the solution explorer and select "Add Reference..."
2. Expand the "Assemblies" list and select the "Extensions" option.
3. From the list, check the box next to each item to select the following four references:
	"Microsoft.Office.Interop.Excel"
	"Microsoft.Office.Interop.Word"
	"Microsoft.Vbe.Interop"
	"office"
4. Click the OK button to close the References dialog.

	Reference Name:	Microsoft.Office.Interop.Excel
	Type:			.NET
	Version:		v14.0.0.0 (or higher)
	File:			Microsoft.Office.Interop.Excel.dll

	Reference Name:	Microsoft.Office.Interop.Word
	Type:			.NET
	Version:		v14.0.0.0 (or higher)
	File:			Microsoft.Office.Interop.Word.dll

	Reference Name:	Microsoft.Vbe.Interop
	Type:			.NET
	Version:		v14.0.0.0 (or higher)
	File:			Microsoft.Vbe.Interop.dll

	Reference Name:	Office
	Type:			.NET
	Version:		v14.0.0.0 (or higher)
	File:			Office.dll

You should now be able to start/compile the application.
Enjoy!