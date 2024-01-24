import requests
import openpyxl

service_url = "https://tiles.arcgis.com/tiles/hUPR9iC6qnMcwWsa/arcgis/rest/services/Lebanon_Basemap/MapServer"
response = requests.get(f"{service_url}?f=json")
service_info = response.json()

layers = service_info["layers"]
tables = service_info["tables"]
json_title = service_info.get("documentInfo", {}).get("Title")

layer_data = []
for layer in layers :
    layer_url = f"{service_url}/{layer['id']}"
    layer_name = layer["name"]
    fields_url = f"{layer_url}?f=json"
    fields_response = requests.get(fields_url)
    fields_info = fields_response.json()
    layer_fields = fields_info["fields"]
    for field in layer_fields:
        field_name = field["name"]
        field_type = field["type"]
        field_alias = field["alias"]
        layer_data.append({"Layer URL": layer_url, "Layer Name": layer_name, "Field Name": field_name,
                           "Field Type": field_type, "Field Alias": field_alias})

table_data = []
for table in tables:
    table_url = f"{service_url}/{table['id']}"
    table_name = table["name"]
    fields_url = f"{table_url}?f=json"
    fields_response = requests.get(fields_url)
    fields_info = fields_response.json()
    table_fields = fields_info["fields"]
    for field in table_fields:
        field_name = field["name"]
        field_type = field["type"]
        field_alias = field["alias"]
        table_data.append({"Table URL": table_url, "Table Name": table_name, "Field Name": field_name,
                           "Field Type": field_type, "Field Alias": field_alias })

filename = f"{json_title or 'untitled'}.xlsx"
workbook = openpyxl.Workbook()

sheet_layers = workbook.active
sheet_layers.title = "Layers"
sheet_layers.append(["Layer URL", "Layer Name", "Field Name", "Field Type", "Field Alias"])
for layer in layer_data:
    sheet_layers.append([layer["Layer URL"], layer["Layer Name"], layer["Field Name"],
                         layer["Field Type"], layer["Field Alias"]])

sheet_tables = workbook.create_sheet("Tables")
sheet_tables.append(["Table URL", "Table Name", "Field Name", "Field Type", "Field Alias"])
for table in table_data:
    sheet_tables.append([table["Table URL"], table["Table Name"], table["Field Name"],
                         table["Field Type"], table["Field Alias"]])

workbook.save(filename)
