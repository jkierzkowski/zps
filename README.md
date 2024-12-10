# Description:
**xlsx_to_json.py** is a script that converts specific types of .xlsx files to JSON.
**template.json** and dict_json_x.json must be in the same folder while running the script.
**template.json** includes a schematic version of the final JSON with data types("str" or "num") for each object value(example: jsonName: "num").
 **dict_json_x.json** is a JSON file containing only objects from template.json with a value from the .xlsx file(example: jsonName: "xlsxColumnName").
