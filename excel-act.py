import pandas as pd
import yaml
import sys
import os

# --- DEFAULT FILES ---
default_input = "./inputs/input.xlsx"
default_output = "./outputs/output.yaml"

# --- COMMAND-LINE ARGUMENTS ---
if len(sys.argv) == 3:
    input_excel = sys.argv[1]
    output_yaml = sys.argv[2]
elif len(sys.argv) == 1:
    input_excel = default_input
    output_yaml = default_output
else:
    print("Usage: python3 script.py <input_excel> <output_yaml>")
    print("Or run without arguments to use defaults: ./inputs/input.xlsx → ./outputs/output.yaml")
    sys.exit(1)

print(f"📘 Using input file: {input_excel}")
print(f"📝 Output will be saved to: {output_yaml}")

# --- ENSURE OUTPUT DIRECTORY EXISTS ---
os.makedirs(os.path.dirname(output_yaml), exist_ok=True)

# --- LOAD EXCEL ---
df = pd.read_excel(input_excel, sheet_name=0, header=None)

# --- FIND SECTIONS ---
sections = {}
for i, val in enumerate(df.iloc[:, 0]):
    if isinstance(val, str):
        v = val.strip().lower()
        if v in ["node_types", "nodes", "links"]:
            sections[v] = i

# --- PARSE NODE TYPES ---
node_types_start = sections["node_types"] + 1
node_headers = df.iloc[node_types_start].dropna().tolist()
node_types_data = df.iloc[node_types_start + 1 : sections["nodes"], :len(node_headers)]
node_types_data.columns = node_headers

node_types_dict = {}
for _, row in node_types_data.iterrows():
    if pd.isna(row[0]):
        continue
    node_type_name = str(row[node_headers[0]])
    details = {
        str(k): str(v)
        for k, v in row.items()
        if k != node_headers[0] and pd.notna(v) and str(v).strip() != ""
    }
    node_types_dict[node_type_name] = details

# --- PARSE NODES ---
nodes_start = sections["nodes"]
node_headers = df.iloc[nodes_start].dropna().tolist()
links_row = sections["links"]
nodes_data = df.iloc[nodes_start + 1 : links_row, :len(node_headers)]
nodes_data.columns = node_headers

nodes_yaml_list = []
for _, row in nodes_data.iterrows():
    if pd.isna(row[0]):
        continue
    node_name = str(row[node_headers[0]])
    node_info = {
        str(k): v
        for k, v in row.items()
        if k != node_headers[0] and pd.notna(v) and str(v).strip() != ""
    }
    nodes_yaml_list.append({node_name: node_info})

# --- PARSE LINKS ---
links_start = sections["links"] + 2  # skip 'links' and 'connection'
links_data = df.iloc[links_start:, :4]

links_yaml_lines = ["links:"]
for _, row in links_data.iterrows():
    if any(pd.isna(row[i]) for i in range(4)):
        continue
    links_yaml_lines.append("    - connection:")
    links_yaml_lines.append(f"        - {row[0]}:{row[1]}")
    links_yaml_lines.append(f"        - {row[2]}:{row[3]}")

links_yaml_str = "\n".join(links_yaml_lines)

# --- BUILD FINAL YAML STRUCTURE ---
yaml_str = yaml.dump({**node_types_dict, "nodes": nodes_yaml_list}, sort_keys=False)
yaml_output = yaml_str.rstrip() + "\n\n" + links_yaml_str + "\n"

# --- WRITE FILE ---
with open(output_yaml, "w") as f:
    f.write(yaml_output)

print(f"✅ YAML file created successfully at: {output_yaml}")
