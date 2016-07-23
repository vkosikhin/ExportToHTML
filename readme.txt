ExportToHTML Microsoft Excel add-in

Current Version: 2.0.1

Valeriy Kosikhin, 2016

vkosikhin@gmail.com

1. Description

This is a tool to create tables in HTML from Excel spreadsheets. Although Excel itself can save spreadsheets as web pages, attributes are appended to HTML tags in order for the result to look like it does in Excel. There is no way to export 'clean' HTML from Excel using the built-in functionality.

What this add-in does is create HTML either with 'clean' tags or tag attributes. The latter can be used to partially preserve cell format from Excel or add whatever custom attributes that are required.

2. Features

	a) The output can be saved to the file or saved and opened immediately in a text editor or web browser, depending on the extension.
	b) The first row of the table is marked as either <th></th> or <td></td>.
	c) It is possible to preserve Excel cell format including text alignment and font style by adding CSS style attributes to the individual cells in a table.
	d) Custom class and style attributes can be added to all instances of a given tag.
	e) Rules can be created to add custom style and class attributes to the <td> or <th> tag of the cell on a condition defined by the cell's format in Excel.
	f) Options are configurable via a user form GUI (see Compatibility).
	g) Text is saved in Unicode.

3. Compatibility

ExportToHTML was developed for Microsoft Office 2016 and tested in Windows 10 and OS X 10.11. Compatibility with the previous Office versions is neither tested nor guaranteed, although the add-in is supposed to work properly in Office 2010/2011 and later.

As user forms support in Office 2016 for Mac is currently not implemented, there is no GUI to make changes to the behavior of the add-in. Options are instead has to be set manually in the configuration file.

In Office 2011 the GUI should, theoretically, work as it does in Windows.

4. Installation

The ExportToHTML.xlam file can be opened directly and stays active for the current Excel session. For the permanent installation use the 'Excel Add-ins' menu in Excel.

On Mac, place the additional file 'WriteToFile.scpt' in the '/Users/Username/Library/Application Scripts/com.microsoft.Excel/' folder. Create the folder if necessary (mind the uppercase in the name).

In Windows, ExportToHTML keeps its settings in the Registry ('\HKEY_CURRENT_USER\SOFTWARE\VB and VBA Program Settings\ExportToHTML'), whereas on Mac they are stored in the file '/Users/Username/Library/Group Containers/UBF8T346G9.Office/VB Settings/ExportToHTML.plist'.

If either the Registry key or the file are not present, the program will create one of these on the first launch.

5. Usage

ExportToHTML creates a custom Ribbon tab with buttons 'Save As', 'Save As And Open' and 'Table Format'.

Select the cells range that you want to be saved as HTML and press 'Save As' or 'Save As And Open'. 'Table Format' will either open the GUI to make changes to the format of the output (Windows) or it will open the configuration file 'ExportToHTML.plist' (Office 2016 for Mac).

6. Changing options in Office 2016 for Mac

The default content of the ExportToHTML.plist is as follows.

<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
	<key>CustomRules\Class</key>
	<dict>
		<key>1</key>
		<string>?</string>
	</dict>
	<key>CustomRules\Condition</key>
	<dict>
		<key>1</key>
		<string> ; ; ; </string>
	</dict>
	<key>CustomRules\Style</key>
	<dict>
		<key>1</key>
		<string>?</string>
	</dict>
	<key>General</key>
	<dict>
		<key>AddStyleToEmptyCells</key>
		<string>False</string>
		<key>EnableCustomFormat</key>
		<string>False</string>
		<key>FirstRowTH</key>
		<string>False</string>
		<key>PreserveFontStyle</key>
		<string>False</string>
		<key>PreserveHorizontalAlignment</key>
		<string>False</string>
		<key>PreserveVerticalAlignment</key>
		<string>False</string>
		<key>tableClass</key>
		<string>?</string>
		<key>tableStyle</key>
		<string>?</string>
		<key>tbodyClass</key>
		<string>?</string>
		<key>tbodyStyle</key>
		<string>?</string>
		<key>tdClass</key>
		<string>?</string>
		<key>tdStyle</key>
		<string>?</string>
		<key>thClass</key>
		<string>?</string>
		<key>thStyle</key>
		<string>?</string>
		<key>trClass</key>
		<string>?</string>
		<key>trStyle</key>
		<string>?</string>
	</dict>
</dict>
</plist>

The keys in the ExportToHTML.plist file which hold 'False' values can be set to 'True' to enable the corresponding options.

The keys with '?' values in the 'General' section contain style and class attributes which will be added to the specific tags in the output file. '?' means that the attribute will not be used. The keys are not allowed to hold empty values in Visual Basic for Applications on Mac.

There is no need to specify attributes as 'style=SLYLE' or 'class=CLASS', as 'style=' and 'class=' will be added automatically.

Sections 'CustomRules\Condition', 'CustomRules\Style' and 'CustomRules\Class' contain rules to add custom tag attributes to the cells on specific conditions. Keys with the same number from the three sections compose a single rule. There must be three keys with the same number for each rule.

The key in the 'CustomRules\Condition' specify the cell format in Excel which will trigger the program to add the style and the class as defined by the keys in the 'CustomRules\Style' and 'CustomRules\Class' sections for that rule.

Values of the keys in the 'CustomRules\Condition' section are formatted as 'FONTSTYLE;FONTSIZE;R,G,B;R,G,B'. The latter values define the font and the fill color of the cell to trigger the rule. Example: 'Bold;11;0,0,0;255,255,255'. Each of the parameters separated by ";" can be left blank, but the separators must remain.