rules = ["CPRE30-C", "PRE31-C", "PRE32-C", "DCL30-C", "DCL36-C", "DCL37-C", "DCL38-C", "DCL39-C", "DCL40-C", "DCL41-C", "CWE-365", "CWE-478", "CWE-484", "CWE-783", "CWE-839", "INT02-C", 
        "INT30-C", "INT31-C", "INT32-C", "INT33-C", "INT34-C", "INT35-C", "INT36-C", "CWE-131", "FLP30-C", "FLP34-C", "FLP36-C", "STR30-C", "STR31-C", "STR32-C", "STR37-C", "STR38-C", "CWE-135",  
        "EXP12-C", "EXP30-C", "EXP33-C", "CWE-Group1", "CWE-481", "CWE-482", "CWE-606", "CWE-674", "EXP34-C", "EXP36-C", "EXP39-C", "EXP40-C", "EXP42-C", "EXP43-C", "EXP44-C", "EXP45-C", "EXP46-C", 
        "EXP47-C", "CWE-785", "ARR30-C", "ARR32-C", "ARR36-C", "ARR37-C", "ARR38-C", "ARR39-C", "CWE-130", "CWE-124", "CWE-806", "CWE-Group2", "CWE-253", "ERR30-C", "ERR33-C", "ERR34-C", "MEM30-C", 
        "MEM31-C", "MEM33-C", "MEM34-C", "MEM35-C", "MEM36-C", "CWE-244", "FIO30-C", "FIO34-C", "FIO37-C", "FIO39-C", "FIO40-C", "FIO41-C", "FIO42-C", "FIO47-C", "CWE-362", "ENV30-C", "ENV31-C", 
        "ENV32-C", "ENV33-C", "SIG30-C", "SIG35-C", "CON30-C", "CON31-C", "CON32-C", "CON33-C", "CON34-C", "CON35-C", "CON36-C", "CON37-C", "CON38-C", "CON39-C", "CON40-C", "MSC30-C", 
        "MSC32-C", "MSC33-C", "MSC37-C", "MSC38-C", "MSC39-C", "CWE-14", "CWE-558", "CWE-Group3", "CWE-320", "CWE-326", "CWE-327", "CWE-Group4", "CWE-353", "CWE-522", "CWE-922", "POS30-C", 
        "POS33-C", "POS34-C", "POS38-C", "POS39-C", "POS44-C", "POS48-C", "POS54-C", "CWE-413", "CWE-489", "CWE-483", "CWE-561"]

rule_counts = {rule: 0 for rule in rules}

with open("Result_List_new.txt", "r", encoding="utf-8") as file:
    for line in file:
        for rule in rules:
            if rule in line:
                if rule == "ARR30-C" and "Same as ARR30-C List" in line:
                    continue
                rule_counts[rule] += 1

# Save to "NumberOfRules.txt"
with open("NumberOfRules.txt", "w", encoding="utf-8") as output:
    for rule, count in rule_counts.items():
        output.write(f"{rule} {count}\n")

print("Summary saved to NumberOfRules.txt")
