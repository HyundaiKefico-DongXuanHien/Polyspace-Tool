# ============================================================================
# FILE NAME     : version5.py
# AUTHOR        : DONG XUAN HIEN
# DIVISION      : HYUNDAI KEFICO Co.,Ltd.
# DESCRIPTION   : Tool automation Polyspace
# HISTORY       : 20/11/2025
# ============================================================================

import os
import time
import pyautogui
import pyperclip
import threading
import keyboard
from charset_normalizer import from_path

#----------------------------------- H I E R A C H Y -----------------------------------------------

data_hierarchy = {
    "Defect": {
        "Data flow": {
            "Dead code": {
                "Status": "No action planned",
                "Severity": "Medium",
                "Comment": "Exception. C-POS-012(CWE-561)"
            },
            "Useless if": {
                "Status": "Unreviewed",
                "Severity": "Unset",
                "Comment": "No subject to the coding guide test."
            },
            "Non-initialized variable": {
                "Status": "Unreviewed",
                "Severity": "Unset",
                "Comment": "No subject to the coding guide test."
            },
            "Write without a further read": {
                "Status": "Not a defect",
                "Severity": "Medium",
                "Comment": "Write variables for class execution"
            },
            "Code deactivated by constant false condition": {
                "Status": "No action planned",
                "Severity": "Medium",
                "Comment": "Exception. C-POS-012(CWE-570)"
            }
        },
        
        "Good practice": {
            "Line with more than one statement": {
                "Status": "Unreviewed",
                "Severity": "Unset",
                "Comment": "No subject to the coding guide test."
            },
            "Incorrectly indented statement": {
                "Status": "Not a defect",
                "Severity": "Medium",
                "Comment": "It is a necessary part for logic execution. There is no problem because it is a code that must be executed once after all if parts have been executed. Also, there seems to be a bug in Polyspace Bug Finder execution. If the above code is a problem, the problem should be caught in the whole c file, but only a specific part is caught. \nAs the module is a manual code, the meaning of the context does not change, and it is not an if-else if-else statement, a switch statement, or a repeat statement, so no braces are required"                           
            },
            "Hard-coded buffer size": {
                "Status": "Unreviewed",
                "Severity": "Unset",
                "Comment": "No subject to the coding guide test."
            },
            "Hard-coded loop boundary": {
                "Status": "Unreviewed",
                "Severity": "Unset",
                "Comment": "No subject to the coding guide test."
            },
            "Bitwise and arithmetic operations on the same data": {
                "Status": "Unreviewed",
                "Severity": "Unset",
                "Comment": "No subject to the coding guide test."
            }
        },
        
        "Numerical": {
            "Unsigned integer overflow": {
                "Status": "Not a defect",
                "Severity": "Unset",
                "Comment": "If there is an error(ex: less than 2) in size, there is no problem because it sends an error message and does not go to the calculation formula"
            },
            "Unsigned integer constant overflow": {
                "Status": "Unreviewed",
                "Severity": "Unset",
                "Comment": "No subject to the coding guide test."
            },
            "Unsigned integer conversion overflow": {
                "Status": "Unreviewed",
                "Severity": "Unset",
                "Comment": "No subject to the coding guide test."
            },
            "Integer constant overflow": {
                "Status": "Unreviewed",
                "Severity": "Unset",
                "Comment": "No subject to the coding guide test."                          
            },
            "Integer overflow": {
                "Status": "Unreviewed",
                "Severity": "Unset",
                "Comment": "No subject to the coding guide test."                          
            },
            "Integer conversion overflow": {
                "Status": "Unreviewed",
                "Severity": "Unset",
                "Comment": "No subject to the coding guide test."                          
            },
            "Shift of a negative value": {
                "Status": "Unreviewed",
                "Severity": "Unset",
                "Comment": "No subject to the coding guide test."                          
            },
            "Right operand of shift operation outside allowed bounds": {
                "Status": "Unreviewed",
                "Severity": "Unset",
                "Comment": "No subject to the coding guide test."                          
            },
            "Precision loss in integer to float conversion": {
                "Status": "Unreviewed",
                "Severity": "Unset",
                "Comment": "No subject to the coding guide test."                          
            },
            "Integer precision exceeded": {
                "Status": "Unreviewed",
                "Severity": "Unset",
                "Comment": "No subject to the coding guide test."                          
            },
            "Sign change integer conversion overflow": {
                "Status": "Unreviewed",
                "Severity": "Unset",
                "Comment": "No subject to the coding guide test."                          
            },
            "Bitwise operation on negative value": {
                "Status": "Unreviewed",
                "Severity": "Unset",
                "Comment": "No subject to the coding guide test."                          
            }
        },
        
        "Programming": {
            "Invalid use of == operator": {
                "Status": "Not a defect",
                "Severity": "Medium",
                "Comment": "It is Auto generated code via ASCET(Not manual code). This code has no incorrect operation"
            },
            "Overlapping assignment": {
                "Status": "Unreviewed",
                "Severity": "Unset",
                "Comment": "No subject to the coding guide test."
                            
            },
            "Qualifier removed in conversion": {
                "Status": "Unreviewed",
                "Severity": "Unset",
                "Comment": "No subject to the coding guide test."                       
            },
            "Floating point comparison with equality operators": {
                "Status": "Unreviewed",
                "Severity": "Unset",
                "Comment": "No subject to the coding guide test."               
            },
            "Misuse of sign-extended character value": {
                "Status": "Unreviewed",
                "Severity": "Unset",
                "Comment": "No subject to the coding guide test."               
            }
        },
        
        "Static memory": {
            "Pointer access out of bounds": {
                "Status": "Not a defect",
                "Severity": "Unset",
                "Comment": "Same as ARR30-C List"
            },
            "Array access out of bounds": {
                "Status": "Not a defect",
                "Severity": "Unset",
                "Comment": "Same as ARR30-C List"            
            },
            "Unreliable cast of pointer": {
                "Status": "Not a defect",
                "Severity": "Low",
                "Comment": "When memory dumps on L9783 IC, VciIf_Type is used as a uint32 structure because it imports 32 bits of memory, but the data inside is less than 8 bits of data.  In order to deliver data for ASW use, the ASW API accesses the data with unit8, so the developer intentionally passes the pointer to 8 bits to prevent errors in memory access  between ASW and BSW."               
            }
        }
    },
    
    
    "SEI CERT C": {
        "ARR": {
            "ARR30-C": {
                "Status": "Not a defect",
                "Severity": "High",
                "Comment": "Input index range is limited by logic and data."
            }
        },
        
        "DCL": {
            "DCL37-C": {
                "Status": "No action planned",
                "Severity": "Low",
                "Comment": "Exception. C-DCI-003(DCL-37)"
            },
            "DCL36-C": {
                "Status": "To fix",
                "Severity": "Medium",
                "Comment": "(temp) Modified SW to be released"
            }
        },
        
        "EXP": {
            "EXP39-C": {
                "Status": "No action planned",
                "Severity": "Medium",
                "Comment": "Exception.C-EXP-011(EXP39-C)"
            },
            "EXP30-C": {
                "Status": "Not a defect",
                "Severity": "Medium",
                "Comment": "CASE-EXP30-C-DefinedAsVolatile. The Rom Data/ARRAY/Non-volatile variable(NV) is declared as volatile variable in the code, but the RomData is fixed data and ARRAY/NV has no risk to changing value while proceed. Refer the \"Prrof Data\" file among the submitted attachments. \n \n자기참조 위반의 경우, 모델 수정 필요"
            },
            "EXP33-C": {
                "Status": "Not a defect",
                "Severity": "High",
                "Comment": "CASE-EXP33-C-NOINIT. \n \n EXP33-C는 신규 검출 되면 안됨."
            },
            "EXP36-C": {
                "Status": "Not a defect",
                "Severity": "Low",
                "Comment": "For this violation is due to the type difference of argument <void *dst> and casted variable <char *d>.  However, the argument <void *dst> was previously aligned (or defined) to refer only to the received CAN buffer.  CAN buffer is strictly formatted by specification (DLC, signal length, signal address, it can not be distorted or changed by any other external disturbance.  Therefore, this code does not affect SW functional actions nor security attacks."
            }
        },
        
        "FLP": {
            "FLP36-C": {
                "Status": "No action planned",
                "Severity": "Low",
                "Comment": "Exception.C-FLP-003(FLP36-C)"
            }
        },
        
        "INT": {
            "INT02-C": {
                "Status": "Not a defect",
                "Severity": "High",
                "Comment": "Operaion ~(tilde) was executed for just negative number. Conversions do not result in lost data \n \nOperation ~(tilde) was executed same in signed/unsigned type. Conversions do not result in lost data."
            },
            "INT30-C": {
                "Status": "Not a defect",
                "Severity": "High",
                "Comment": "If there is an error in size, there is no problem because it sends an error message and does not go to the calculation formula \n \nThe values are intended value inserted by calibration engineer who know about the overflow. And it is ROM data. So no problems because of it is ROM data,  the value is supervised. Also the data is written with data set"
            },
            "INT31-C": {
                "Status": "Not a defect",
                "Severity": "High",
                "Comment": "CASE-INT31-C-MISINTERPRETED. Dead code by cal data. \n \nOperaion ~(tilde) was executed for just negative number. Conversions(from int16 to unsigned int32) do not result in lost data. \n \nOperation ~(tilde) was executed same in signed/unsigned type. Conversions do not result in lost data. \n \nOperation <<(Shift left) was executed same in signed/unsigned type. Conversions do not result in lost data \n \nConversion to the data type of the base bit-field. Conversions do not result in lost data \n \nThis value is only for diagnostic device and not used for input or calculation to other modules. \n \nCASE-INT31-C-MISINTERPRETED. Intended operatopmn for formula changing \n \nOperation ^(Exclusive OR) was executed same in signed/unsigned type. Conversions do not result in lost data. \n \nOperation -(Minus) was executed for just negative number. Then conversion to unsigned type was adapted"
            },
            "INT32-C": {
                "Status": "Not a defect",
                "Severity": "High",
                "Comment": "Overflow would be checked with below if condition \n \nOverflow would be checked with upper high part condition"
            },
            "INT34-C": {
                "Status": "Not a defect",
                "Severity": "High",
                "Comment": "CASE-INT34-C. After type conversion, shift operation is performed by the constant specified within the variable range, so no problem \n \nCASE-INT34-C. During auto-code generation in ASCET, the logic that matches the implementation of variables used for calculation is operating,  and if there is a possibility of overflow, automatic prevention of overflow is automatically generated.  Therefore, in this case, there is no problem in logic, since the sint8 type variable (positive number, 0 to 127) is converted to  unit16 and left shift operation is performed by 2 smaller than the number of the upper 8 bits padded with 0."
            },
            "INT35-C": {
                "Status": "Not a defect",
                "Severity": "High",
                "Comment": "at line 2146, 2147 in the auto generated code(GPFltRD), _t1sint16 is GPFlt_nrSotTmp-1 and it limted bt GFPltR_nrSotAry_C. So there is no problem because it is within the bit number(0 to7) of the corresponding variable."
            }
        },
        
        "MEM": {
            "MEM35-C": {
                "Status": "Not a defect",
                "Severity": "High",
                "Comment": "Automatically generated code for post-processing, reflecting code and data and considering that it does not deviate from the specified size \n \nThe values are intended value inserted by calibration engineer who know about the overflow. And it is ROM data. So no problems Because of it is ROM data, the value is supervised. Also the data is written with data set"
            }
        },
        
        "MSC": {
            "MSC37-C": {
                "Status": "Not a defect",
                "Severity": "High",
                "Comment": "The code runs until the condition of \"if\" is satisfied."
            }
        }
    }
}

status_hierarchy = {
    "Unreviewed": 0,
    "To investigate": 1,
    "To fix": 2,
    "Justified": 3,
    "No action planned": 4,
    "Not a defect": 5,
    "Other": 6,
    }
    
severity_hierarchy = {
    "Unset": 0,
    "High": 1,
    "Medium": 2,
    "Low": 3,
    }    

P_status = (1525, 190)   #(3198, 157)
P_severity = (1527, 225) #(3201,188)
P_comment = (1736, 194) # (3351, 158)

P = (825, 132)

level1_compare = ""
level2_compare = ""

stop_flag = False
#----------------------------------- F I N D   I N F O R M A T I O N ----------------------------
def find_group_info(data, level3_name):
    for level1_name, level2_groups in data.items():
        for level2_name, level3_groups in level2_groups.items():
            for level3_key, info in level3_groups.items():
                if level3_key.lower() == level3_name.lower():
                    return {
                        "Level1": level1_name,
                        "Level2": level2_name,
                        "Level3": level3_key,
                        "Status": info.get("Status"),
                        "Severity": info.get("Severity"),
                        "Comment": info.get("Comment")
                    }
    return None


#----------------------------------- H A N D L E   L E V E L 3 ----------------------------
# Hàm xử lý riêng cho phần tử đầu tiên
def handle_first_item(result, P):
    #--------Click Level 2--------------------
    time.sleep(5)
    keyboard.send("add")
    #--------Click Level 3--------------------
    time.sleep(2)
    pyautogui.press("down")
    time.sleep(2)
    keyboard.send("add")
    
    # Click properties
    time.sleep(2)
    pyautogui.press("down")
    time.sleep(1)
    pyautogui.keyDown('ctrl')
    pyautogui.press('a')
    pyautogui.keyUp('ctrl')
    
    #------------------ Handle Status ------------------
    status_text = result.get("Status", "")
    status_offset = status_hierarchy.get(status_text, 0)  # default 0 nếu không tìm thấy
    status_offset = int(status_offset)

    # Click vào ô Status
    time.sleep(2)
    pyautogui.moveTo(P_status)
    pyautogui.click()
    time.sleep(2)

    # Click vào lựa chọn Status tương ứng
    for _ in range(status_offset):
        pyautogui.press("down")
        time.sleep(1)
    pyautogui.press("Enter")

    #------------------ Handle Severity ------------------
    severity_text = result.get("Severity", "")
    severity_offset = severity_hierarchy.get(severity_text, 0)
    severity_offset = int(severity_offset)

    # Click vào ô Severity
    time.sleep(2)
    pyautogui.moveTo(P_severity)
    pyautogui.click()
    time.sleep(2)

    # Click vào lựa chọn Severity tương ứng
    for _ in range(severity_offset):
        pyautogui.press("down")
        time.sleep(1)
    pyautogui.press("Enter")   
    
    #------------------ Handle Comment ------------------
    # Click vào ô Comment
    time.sleep(2)
    pyautogui.moveTo(P_comment)
    pyautogui.click()    
    
    time.sleep(2)
    comment_text = result.get("Comment", "")
    pyperclip.copy(comment_text)
    pyautogui.hotkey("ctrl", "v")
    
    #----------------Click default_area --------------------------
    time.sleep(2)
    pyautogui.moveTo(P)
    pyautogui.click()
    time.sleep(1)
    pyautogui.press("up")
    time.sleep(1)
    keyboard.send("subtract")
   
def handle_result(result, P):
    #--------Click Level 2--------------------
    time.sleep(5)
    pyautogui.press('down')
    time.sleep(1)
    keyboard.send("add")
    
    # Click properties
    time.sleep(2)
    pyautogui.press("down")
    time.sleep(1)
    pyautogui.keyDown('ctrl')
    pyautogui.press('a')
    pyautogui.keyUp('ctrl')
    
    #------------------ Handle Status ------------------
    status_text = result.get("Status", "")
    status_offset = status_hierarchy.get(status_text, 0)  # default 0 nếu không tìm thấy
    status_offset = int(status_offset)

    # Click vào ô Status
    time.sleep(2)
    pyautogui.moveTo(P_status)
    pyautogui.click()
    time.sleep(2)

    # Click vào lựa chọn Status tương ứng
    for _ in range(status_offset):
        pyautogui.press("down")
        time.sleep(1)
    pyautogui.press("Enter")

    #------------------ Handle Severity ------------------
    severity_text = result.get("Severity", "")
    severity_offset = severity_hierarchy.get(severity_text, 0)
    severity_offset = int(severity_offset)

    # Click vào ô Severity
    time.sleep(2)
    pyautogui.moveTo(P_severity)
    pyautogui.click()
    time.sleep(2)

    # Click vào lựa chọn Severity tương ứng
    for _ in range(severity_offset):
        pyautogui.press("down")
        time.sleep(1)
    pyautogui.press("Enter")   
    
    #------------------ Handle Comment ------------------
    # Click vào ô Comment
    time.sleep(2)
    pyautogui.moveTo(P_comment)
    pyautogui.click()    
    
    time.sleep(2)
    comment_text = result.get("Comment", "")
    pyperclip.copy(comment_text)
    pyautogui.hotkey("ctrl", "v")
    
    #----------------Click default_area --------------------------
    time.sleep(2)
    pyautogui.moveTo(P)
    pyautogui.click()
    time.sleep(1)
    pyautogui.press("up")
    time.sleep(1)
    keyboard.send("subtract")


#----------------------------------- S U B   P R O G R A M ----------------------------
# Thread giám sát phím ESC
def monitor_keyboard():
    global stop_flag
    print("🎯 Đang giám sát phím ESC để dừng chương trình...")
    while True:
        if keyboard.is_pressed("esc"):
            print("🛑 Phát hiện phím ESC! Dừng chương trình khẩn cấp.")
            stop_flag = True
            os._exit(1)  # Dừng toàn bộ chương trình ngay lập tức
        time.sleep(0.1)


#----------------------------------- M A I N   P R O G R A M ----------------------------
# Gọi thread ngay khi khởi chạy chương trình
keyboard_thread = threading.Thread(target=monitor_keyboard, daemon=True)
keyboard_thread.start()

# Vòng lặp nhập từ người dùng
while True:
    user_input = input("Nhập danh sách tên nhóm cấp 3 (cách nhau bởi dấu ',') hoặc 'exit': ").strip()
    if user_input.lower() == "exit":
        break

    level3_names = [name.strip() for name in user_input.split(",") if name.strip()]
    
    for i, level3_name in enumerate(level3_names):
        result = find_group_info(data_hierarchy, level3_name)
        if result:
            if i == 0:
                time.sleep(3)
                level1_compare = result["Level1"]
                level2_compare = result["Level2"]
                handle_first_item(result, P)
            elif level1_compare == result["Level1"] and level2_compare == result["Level2"]:
                handle_result(result, P)
            elif level1_compare == result["Level1"] and level2_compare != result["Level2"]:
                pyautogui.press("down")
                level2_compare = result["Level2"]
                handle_first_item(result, P)      
            else:
                time.sleep(1)
                pyautogui.press("down")
                time.sleep(1)
                pyautogui.press("down")
                level1_compare = result["Level1"]
                level2_compare = result["Level2"]
                handle_first_item(result, P)
        else:
            print(f"⚠️ Không tìm thấy nhóm cấp 3: {level3_name}")