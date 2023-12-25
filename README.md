# TBI_dict_excel_to_word
Python script to compile excel file into docx. TG bot too?

Windows Exe file compiled by pyinstaller. Maybe should add some settings for styling
Recompile exe with --onefile and --windowed options. 


doc-to-exl:
Takes initial pattern "pfn. utvj" to get terminology[beginning of line:pattern] separated.
Then uses manually positioned "/*" symbols to divide definition from example
Then it is trimmed with book reference pattern. Hard coded one. All of them are for now.
Regex is used to fetch page info in reference part.

