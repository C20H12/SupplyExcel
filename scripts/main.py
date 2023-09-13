import json
import uuid

# with open("./sizing_map.json") as f:
#   data = f.read()
#   converted_data = json.loads(data)

def get_id():
  return str(uuid.uuid4())[:6]

def convert(d, collection_name, result: list, nsn_map=None):
  result.append(f"Dim {collection_name} As Collection")
  result.append(f"Set {collection_name} = New Collection")

  if any(map(lambda key: key[0].isdigit(), d.keys())):
    keys_id = get_id()
    result.append(f"Dim Keys_{keys_id}() As Variant")
    sorted_keys = sorted(d.keys(), key=lambda k: float(k))
    result.append(f"Keys_{keys_id} = Array({','.join(sorted_keys)})")
    result.append(f"{collection_name}.Add Keys_{keys_id}, \"Keys\"")

  for key, value in d.items():
    key_modified = key.replace(".", "_") + "_" + get_id()
    if type(value) is dict:
      convert(value, f"Value_{key_modified}", result, nsn_map)
      result.append(f"{collection_name}.Add Value_{key_modified}, \"{key}\"")
    elif type(value) is list:
      result.append(f"Dim Value_{key_modified}() As String")
      joined_arr = ','.join(value)
      result.append(f"Value_{key_modified} = Split(\"{joined_arr}\", \",\")")
      result.append(f"{collection_name}.Add Value_{key_modified}, \"{key}\"")
    else:
      if nsn_map is not None:
        result.append(f"{collection_name}.Add \"{value}==={nsn_map[value]}\", \"{key}\"")
      else:
        result.append(f"{collection_name}.Add \"{value}\", \"{key}\"")
        
  
  return result

loaded = json.load(open("./maps/sizing_map.json"))

nsn_loaded = json.load(open("./maps/nsn_map.json"))

for key, value in loaded.items():
  # print(key,value)
  print(f"Private Function GetSizingData_{key}() As Collection")
  res = convert(value, f"SizingMap_{key}", [], nsn_map=nsn_loaded[key])
  print("    ", end="")
  print(*res, sep="\n    ")
  print(f"    Set GetSizingData_{key} = SizingMap_{key}")
  print("End Function")

  print()