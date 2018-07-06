# JSON-Formatter
Can format JSON/Line files into text files and or excel spreadsheets.
This program uses the Regex, Xlsxwriter, OS packages.
You can install them very easily by typing:
'pip install re' for the Regex package;
'pip install xlsxwriter' for the xlsxwriter package.

Although you're likely not going to have to do this, except maybe for the xlsxwriter package.

Currently this formatter is limited to data formatted like this:

{"name": ["Forensic Safety Group"], "phone": ["(215) 659-2400"], "address": ["543 Davisville Rd", "Willow Grove", ", ", "PA", " ", "19090"]}

However, I might change/update it to be able to format other formatted JSONL files.
