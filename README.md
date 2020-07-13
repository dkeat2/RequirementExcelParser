# Instructions

## Installation
In a command prompt, run `pip install -r requirements.txt` to install the `openpyxl` dependency

## Usage
In a command prompt, run `python main.py FILENAME.xlsx` to parse a workbook. Please see `TestReq.xlsx` for an example of workbook formatting.

The parser will extract requirement data from each sheet. The user will be alerted if any requirements have multiple parents. Ex:
```
The following requirements have multiple parents:
	req-id-01
	req-id-02
```
For now, this is simply to mark requirements with multiple parents for manual investigation.

Next, for requirements with multiple texts, the user will be prompted to select the text they would like to record. Ex:
```
The following requirements have multiple texts. Please select which text to use.

	req-id-01:
		1: Requirement 1 text version 1 
		2: Requirement 1 text version 2 

		Selected Text Number: __
``` 
To select an option, enter the number that appears on the left-hand side of the text.

Finally, the user will be prompted to enter a filename to use for a csv export. 
```
Chose a filename for export: __
```
This can be any string. An output file will be saved in the script's root directory with the following format:
```
id, text, parents, children
req-id-01, Requirement 1 text, req-parent-1;req-parent-2, req-child-1;req-child2
...

```
