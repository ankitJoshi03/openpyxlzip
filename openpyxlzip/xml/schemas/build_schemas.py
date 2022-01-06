#!/usr/bin/env python3

import os
import json
import xmlschema

all_tag_pairs = {}

def parse_type(complex_type):
    element_order = {}
    attributes = {}
    name = complex_type.name
    for attr_key in complex_type.attributes:
        attributes[attr_key] = str(complex_type.attributes[attr_key].type.name)
    if complex_type.has_complex_content():
        for item in complex_type.content.iter_elements():
            if type(item) is xmlschema.XsdElement:
                # print(name, item.name, item.type.name)
                # if item.name in all_tag_pairs and item.type.name not in all_tag_pairs[item.name]:
                #     print("Conflict", item.name, all_tag_pairs[item.name], item.type.name)
                if item.name not in all_tag_pairs:
                    all_tag_pairs[item.name] = []
                if item.type.name not in all_tag_pairs[item.name]:
                    all_tag_pairs[item.name].append(item.type.name)
                element_order[len(element_order)] = (item.name, item.type.name)
            elif type(item) is xmlschema.validators.XsdAnyElement:
                element_order[len(element_order)] = None
            # else:
            #     print("Unhandled", xsd_component.name, item.name, item, type(item))
    return attributes, element_order

all_definitions = {}
all_types = set()
all_namespaces_prefixes = {}
root_namespaces = []
global_attr = {}
global_elem = {}
data_path = os.path.dirname(os.path.realpath(__file__))
parent_path = os.path.dirname(os.path.normpath(data_path))
for filename in os.listdir(data_path):
    if not filename.endswith(".xsd"):
        continue
    schema = xmlschema.XMLSchema(filename, validation='skip')
    for xsd_component in schema.iter_components():
        if type(xsd_component) is xmlschema.validators.XsdComplexType:
            name = xsd_component.name
            attributes, element_order = parse_type(xsd_component)
            # if name in all_definitions:
            #     print("Duplicate", xsd_component)
            all_definitions[name] = {"attributes": attributes, "element_order": element_order}
        elif type(xsd_component) not in all_types:
            # print(type(xsd_component))
            all_types.add(type(xsd_component))
    #Global attributes
    for attr_name in schema.attributes:
        xsd_component = schema.attributes[attr_name]
        global_attr[element_name] = xsd_component.type.name
    #Global elements
    for element_name in schema.elements:
        xsd_component = schema.elements[element_name]
        namespace = schema.target_namespace
        namespaced_name = "{%s}%s" % (namespace, element_name)
        # if namespaced_name in global_elem and xsd_component.type.name != global_elem[namespaced_name]:
        #     print("Overlap", namespaced_name)
        if namespaced_name not in global_elem:
            global_elem[namespaced_name] = xsd_component.type.name
        for prefix in xsd_component.namespaces:
            if prefix == "":
                if xsd_component.namespaces[prefix] not in root_namespaces:
                    root_namespaces.append(xsd_component.namespaces[prefix])
            else:
                # if prefix in all_namespaces_prefixes and all_namespaces_prefixes[prefix] != xsd_component.namespaces[prefix]:
                #     print("Prefix mismatch", prefix, all_namespaces_prefixes[prefix], xsd_component.namespaces[prefix])
                all_namespaces_prefixes[prefix] = xsd_component.namespaces[prefix]

no_conflicts = 0
for tag in all_tag_pairs:
    if len(all_tag_pairs[tag]) == 1:
        no_conflicts += 1
    # else:
    #     print("*****************", tag, all_tag_pairs[tag])

# print(len(all_tag_pairs), no_conflicts)

with open(os.path.join(data_path, "all_definitions.json"), "w") as outfile:
    json.dump(all_definitions, outfile, indent=4)

with open(os.path.join(data_path, "all_namespaces_prefixes.json"), "w") as outfile:
    json.dump(all_namespaces_prefixes, outfile, indent=4)

with open(os.path.join(data_path, "root_namespaces.json"), "w") as outfile:
    json.dump(root_namespaces, outfile, indent=4)

with open(os.path.join(data_path, "root_attrs.json"), "w") as outfile:
    json.dump(global_attr, outfile, indent=4)

with open(os.path.join(data_path, "root_elems.json"), "w") as outfile:
    json.dump(global_elem, outfile, indent=4)

with open(os.path.join(parent_path, "schema.py"), "w") as outfile:
    outfile.write("import json\n\n\n")
    outfile.write("ALL_DEFINITIONS = json.loads(\"\"\"\n")
    outfile.write(json.dumps(all_definitions, indent=4))
    outfile.write("\"\"\")\n")
    outfile.write("ALL_NAMESPACES_PREFIXES = json.loads(\"\"\"\n")
    outfile.write(json.dumps(all_namespaces_prefixes, indent=4))
    outfile.write("\"\"\")\n")
    outfile.write("ROOT_NAMESPACES = json.loads(\"\"\"\n")
    outfile.write(json.dumps(root_namespaces, indent=4))
    outfile.write("\"\"\")\n")
    outfile.write("ROOT_ATTRS = json.loads(\"\"\"\n")
    outfile.write(json.dumps(global_attr, indent=4))
    outfile.write("\"\"\")\n")
    outfile.write("ROOT_ELEMS = json.loads(\"\"\"\n")
    outfile.write(json.dumps(global_elem, indent=4))
    outfile.write("\"\"\")\n")
