# WBClean_XUM

WBClean_XUM is a small Python utility class for cleaning Excel workbooks by **transposing entire sheets**, removing noisy/empty columns, optionally mapping “important” columns using an LLM (Groq/OpenAI-compatible API), converting `.xls → .xlsx` (Windows), and returning the cleaned output as a **pandas DataFrame** or an Excel file.

---

## Installation

### From PyPI
```bash
pip install WBClean_XUM
```

### From GitHub
```bash
pip install WBClean_XUM
```

## Requirements

- Python 3.9+
- Dependencies:
  - `pandas`
  - `openpyxl`
  - `requests`
  - `pywin32` (Windows only; required only for `.xls` conversion)

> **Note:** `.xls` conversion uses COM automation (`win32com`) and requires Microsoft Excel installed on Windows.

---
## Quick Start
1) Clean a raw Excel file and get a DataFrame

```python
import re
import WBClean_XUM

tool = WBClean_XUM.new()


df = tool.XUM_Clean(
    filePath,                    # File path
    pattern=pattern,             # Pattern must be lowercase column name that you are looking in your sheet. if mupltiple present, seperate by |. STR format
    returnDF=True,               # By default returns Dataframe. if toggled to False, will save an .xlsx file in your folder 
    remove_none=True,            # by default will remove all the None rows. Toggle to False if required.
    sheetName=None,              # optional if destinationSheetName specified
    destinationSheet="WBClean_XUM", # Destination sheet name
    getImpFeatures=False,        # By default is set to False, because this requires a groq account and api key of yours
    prompt_ReqFeildString="""
    - col_id_1: What keywords to look for here. Have each seperated by / . 
    - col_id_2: Description / Note (give description only if present or else use Notes)
    - col_id_3: Quantity (QTY / Qty / Order Qty)
    - col_id_4: OEM Part Number / Catalog Number / X ref / Alt Part if explicitly present (not item code), else null
    - add more as per your requirement...
    """,                         # If getImpFeatures == True, then this parameter is required. This is the rules for your LLM.
    prompt_ReqJSONOutputString="""
    {{
      "col_id_1": "enter a valid datatype (int,str,float) and null for fallback",
      "col_id_2": "int or null",
      "col_id_3": "str or null",
      "col_id_4": "float or null",
    }}
    """,                        # If getImpFeatures == True, then this parameter is required. This is the output format specifier for your LLM.
    Key="Groq API Key",          # If getImpFeatures == True, then this parameter is required. This is the output format specifier for your LLM.
    APIUrl="https://api.groq.com/openai/v1/chat/completions", # change as per your requirement
    groqModel="llama-3.3-70b-versatile", # change as per your requirement
    contentType="application/json", # change as per your requirement
    temperature=0, # change as per your requirement
    maxTokens=512 # change as per your requirement
)

print(df.head())
```

2) Use Utilities seperately
### Transpose
```python
XUM_TransposeSheet(
    srcPath,                              # File path
    destinationSheetName="WBClean_XUM",   # Destination sheet name (Recommended)
    srcSheetName=None                     # optional if destinationSheetName specified
)

#TODO: Returns -> 2-D Array
```

### LLM Feature
```python
XUM_LLMFormat(
    prompt_ReqFeildString="""
    - col_id_1: What keywords to look for here. Have each seperated by / . 
    - col_id_2: Description / Note (give description only if present or else use Notes)
    - col_id_3: Quantity (QTY / Qty / Order Qty)
    - col_id_4: OEM Part Number / Catalog Number / X ref / Alt Part if explicitly present (not item code), else null
    - add more as per your requirement...
    """,                         # If getImpFeatures == True, then this parameter is required. This is the rules for your LLM.
    prompt_ReqJSONOutputString="""
    {{
      "col_id_1": "enter a valid datatype (int,str,float) and null for fallback",
      "col_id_2": "int or null",
      "col_id_3": "str or null",
      "col_id_4": "float or null",
    }}
    """,                        # If getImpFeatures == True, then this parameter is required. This is the output format specifier for your LLM.
    Key="Groq API Key",          # If getImpFeatures == True, then this parameter is required. This is the output format specifier for your LLM.
    APIUrl="https://api.groq.com/openai/v1/chat/completions", # change as per your requirement
    groqModel="llama-3.3-70b-versatile", # change as per your requirement
    contentType="application/json", # change as per your requirement
    temperature=0, # change as per your requirement
    maxTokens=512 # change as per your requirement
    prompt_SampleData, # enter "" if nothing to provide else pass-in any array
    prompt_FullCustom=False, # This is a the point that changes everything else. If this is toggled to True, you don't have to specify [prompt_ReqFeildString,prompt_ReqJSONOutputString]
    prompt_Full=None # if prompt_FullCustom==True, then provide your custom written prompt to pass the LLM
)

#TODO: Returns -> json in string format ("{json contents}")
```

### Delete Rows
```python
XUM_DeleteRows(
    Table2dArray=[[],[],[],..], # 2-D array of the table/sheet
    rowIndexList=[0,1,2,..]
)

#TODO: Returns -> table as 2-D Array
```

### Delete Columns
```python
XUM_DeleteColumns(
    Table2dArray=[[],[],[],..],  # 2-D array of the table/sheet
    colIndexList=[0,1,2,..]
)

#TODO: Returns -> table as 2-D Array
```

### Regex based pattern search
```python
XUM_TextPresenceRegex(
    x="Text: Anything",
    pattern=re.compile('regexpattern', re.I)
)

#TODO: Returns -> (Boolean) True or False
```

### XLS Conversion
```python
XUM_XLSConversion(
    xlsPath="xls file path"
)

#TODO: Returns -> xlsx file path
```

For suggesting changes or updates or corrections, please contact: nmkrishnan108@gmail.com