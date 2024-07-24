import os, math
from collections import defaultdict
from PyQt5.QtWidgets import QMessageBox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

MSG_SUCCESS = 0
MSG_WARNING = 1
MSG_ERROR = -1

scores = {}
def getGoogleSheetService():
    # If modifying these scopes, delete the file token.json.
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    
    credential = None
    if os.path.exists('assets/utils/token.json'):
        credential = Credentials.from_authorized_user_file('assets/utils/token.json', SCOPES)
    if not credential or not credential.valid:
        if credential and credential.expired and credential.refresh_token:
            credential.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('assets/utils/credentials.json', SCOPES)
            credential = flow.run_local_server(port=0)
        with open('assets/utils/token.json', 'w') as token:
            token.write(credential.to_json())
    try:
        service = build("sheets", "v4", credentials=credential)
        return service
    except HttpError as err:
        print(err)
        return None
    
def load_spreadsheet_data(wsheetId, msheetId):
    worksheet = getGoogleSheetService().spreadsheets()
    sheet_names = []
    try:
        wsheet_metadata = worksheet.get(spreadsheetId=wsheetId).execute()
        for sheet in wsheet_metadata['sheets']:
            sheet_names.append(sheet['properties']['title'])
        try:
            worksheet.get(spreadsheetId=msheetId).execute()
            return {"status": MSG_SUCCESS, "msg": "Success", "data": sheet_names}
        except:
            return {"status": MSG_ERROR, "msg": "The Google Sheet Service is not able to use for now. Try again later."}
    except:
        return {"status": MSG_ERROR, "msg": "The Google Sheet Service is not able to use for now. Try again later."}
    
def getChromeDriver():
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--log-level=3")
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chrome_options)

    driver.execute_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source":
            "const newProto = navigator.__proto__;"
            "delete newProto.webdriver;"
            "navigator.__proto__ = newProto;"
    })
    return driver
    
def getSheetColumnLabels(start_index, n):
    column_labels = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"]
    sheet_column_labels = []

    for i in range(start_index, n):
        if i < len(column_labels):
            sheet_column_labels.append(column_labels[i])
        else:
            # If you need more than 26 columns, you can extend the labels with combinations like AA, AB, etc.
            div, mod = divmod(i, len(column_labels))
            label = column_labels[mod]
            while div > 0:
                div, mod = divmod(div - 1, len(column_labels))
                label = column_labels[mod] + label
            sheet_column_labels.append(label)

    return sheet_column_labels
        
def getColumnLabelByIndex(ind):
    labels = getSheetColumnLabels(0, 50)
    return labels[ind]

def getProjectPath():
    return os.getcwd().replace("\\", "/")

def getTextValue(list, index):
    try:
        return list[index].find_element(By.CSS_SELECTOR, "div.block-name").get_attribute("title")
    except:
        return ""

def getPedigreeDataFromTable(table):
    result = []
    ## Extract the values will be input in spreadsheet ##
    ids = ["MMM", "MMMM", "MM", "MMF", "MMFM", "M", "MFM", "MFMM", "MF", "MFF", "MFFM", "FMM", "FMMM", "FM", "FMF", "FMFM", "F", "FFM", "FFMM", "FF", "FFF", "FFFM"]
    for id in ids:
        td_elem = table.find_element(By.CSS_SELECTOR, f"td#{id}")
        next_td_elem = td_elem.find_element(By.XPATH, "./following-sibling::td")
        if next_td_elem.get_attribute("class") == "pedigree-cell-highlight":
            label = td_elem.find_element(By.CSS_SELECTOR, "div.block-name").get_attribute("title").title()
            label += "*"
            result.append(label)
        else:
            result.append(td_elem.find_element(By.CSS_SELECTOR, "div.block-name").get_attribute("title").title())
        

    return result

def getLetterGradeBy(g_sire, g_damssire, g_damssire2, g_damssire3):
    letter_grade_constants = {"A+": 5, "A": 4, "A-": 3, "B": 2, "B-": 1}
    sum_grades = (letter_grade_constants[g_sire] if g_sire != None else 0) + (letter_grade_constants[g_damssire] if g_damssire != None else 0) + (letter_grade_constants[g_damssire2] if g_damssire2 != None else 0) + letter_grade_constants[g_damssire3]
    avg_grade_value = get2DigitsFloatValue(float(sum_grades / 4))
    final_grade_value = math.ceil(avg_grade_value)
    if final_grade_value >= 5:
        return {"letter": "A+", "color_info": [36, 246, 0]}
    elif final_grade_value == 4:
        return {"letter": "A", "color_info": [152, 245, 0]}
    elif final_grade_value == 3:
        return {"letter": "A-", "color_info": [245, 246, 0]}
    elif final_grade_value == 2:
        return {"letter": "B", "color_info": [237, 158, 0]}
    elif final_grade_value == 1:
        return {"letter": "B-", "color_info": [255, 1, 1]}
    
def get2DigitsStringValue(input):
    return '%.2f' % float(input)

def get2DigitsFloatValue(input):
    return float('%.2f' % float(input))

def getPositionByPercent(fval, sval, percent):
    diff = sval - fval
    result = int((percent / 100) * diff)
    return fval + result

def groupBySireAndCountHorse(init_data, genType):
    result_dict = defaultdict(lambda: defaultdict(int))
    
    for h, s in init_data:
        result_dict[s.title()][h.title()] += 1
    result_array = []
    for s, h_cnts in result_dict.items():
        for h, cnt in h_cnts.items():
            result_array.append([h, s, str(cnt), ""])
    result_array = sortByIndex(result_array, 2)
    if genType == 0:
        return result_array[:10]
    else:
        return result_array

def sortByRate(arr, genType):
    sorted_arr = sorted(arr, key=lambda x: custom_key(x, 1), reverse=True)
    if genType == 0:
        cutted_arr = sorted_arr[:10]
        last_element = cutted_arr[-1]
        for v in sorted_arr[10:]:
            if v[1] == last_element[1]:
                cutted_arr.append(v)
            else: break
        return cutted_arr
    else:
        return sorted_arr

def sortByCoi(arr, genType):
    sorted_arr = sorted(arr, key=lambda x: float(x[4][:-1]), reverse=True)
    if genType == 0:
        cutted_arr = sorted_arr[:10]
        last_element = cutted_arr[-1]
        for v in sorted_arr[10:]:
            if v[4] == last_element[4]:
                cutted_arr.append(v)
            else: break
        filtered_arr = [x for x in cutted_arr if x[4].strip() != "0.00%"]
        if len(filtered_arr) == 0:
            return sortByVariant(cutted_arr, genType)
        else:
            return filtered_arr
    else:
        return sorted_arr
    
def sortByCoiForUnrated(arr, genType):
    sorted_arr = sorted(arr, key=lambda x: float(x[4][:-1]), reverse=True)
    if genType == 0:
        cutted_arr = sorted_arr[:10]
        last_element = cutted_arr[-1]
        for v in sorted_arr[10:]:
            if v[4] == last_element[4]:
                cutted_arr.append(v)
            else: break
        filtered_arr = [x for x in cutted_arr if x[4].strip() != "0.00%"]
        if len(filtered_arr) == 0:
            return sortByVariant(cutted_arr, genType)
        else:
            return cutted_arr
    else:
        return sorted_arr

def sortByVariant(arr, genType):
    sorted_arr = sorted(arr, key=lambda x: custom_key(x, 2), reverse=True)
    if genType == 0:
        cutted_arr = sorted_arr[:10]
        last_element = cutted_arr[-1]
        for v in sorted_arr[10:]:
            if v[2] == last_element[2]:
                cutted_arr.append(v)
            else: break
        return cutted_arr
    else:
        return sorted_arr
    
def sortByVariant2(arr, genType):
    sorted_arr = sorted(arr, key=lambda x: custom_key(x, 3), reverse=True)
    if genType == 0:
        cutted_arr = sorted_arr[:10]
        last_element = cutted_arr[-1]
        for v in sorted_arr[10:]:
            if v[3] == last_element[3]:
                cutted_arr.append(v)
            else: break
        return cutted_arr
    else:
        return sorted_arr
    
def sortByIndex(arr, ind):
    sorted_arr = sorted(arr, key=lambda x: float(x[ind]), reverse=True)
    cutted_arr = sorted_arr[:10]
    last_element = cutted_arr[-1]
    for v in sorted_arr[10:]:
        if v[ind] == last_element[ind]:
            cutted_arr.append(v)
        else: break
    return cutted_arr

def sortByOtherTiers(arr, genType):
    sorted_arr = sorted(arr, key=lambda x: len(x[0].replace(" ","")), reverse=True)
    sorted_arr = sorted(sorted_arr, key=lambda x: custom_key(x, 3), reverse=True)
    if genType == 0:
        cutted_arr = sorted_arr[:10]
        last_element = cutted_arr[-1]
        for v in sorted_arr[10:]:
            if v[3] == last_element[3]:
                cutted_arr.append(v)
            else: break
        return cutted_arr
    else:
        return sorted_arr

def custom_key(item, ind):
    if item[ind] == "N/A":
        return float('0')
    elif item[ind] == "":
        return float('0')
    else:
        return float(item[ind].replace("%",""))
    
def getJsonDataOfStallion(data):
    jsonObj = {}
    jsonObj["name"] = data[0]

    jsonObj["s"] = {}
    jsonObj["s"]["name"] = data[1]
    jsonObj["s"]["s"] = {}
    jsonObj["s"]["s"]["name"] = data[3]
    jsonObj["s"]["s"]["s"] = {}
    jsonObj["s"]["s"]["s"]["name"] = data[7]
    jsonObj["s"]["s"]["d"] = {}
    jsonObj["s"]["s"]["d"]["name"] = data[8]
    jsonObj["s"]["s"]["d"]["s"] = {}
    jsonObj["s"]["s"]["d"]["s"]["name"] = data[17]
    jsonObj["s"]["s"]["d"]["d"] = {}
    jsonObj["s"]["s"]["d"]["d"]["name"] = data[18]
    jsonObj["s"]["s"]["s"]["s"] = {}
    jsonObj["s"]["s"]["s"]["s"]["name"] = data[15]
    jsonObj["s"]["s"]["s"]["d"] = {}
    jsonObj["s"]["s"]["s"]["d"]["name"] = data[16]
    jsonObj["s"]["d"] = {}
    jsonObj["s"]["d"]["name"] = data[4]
    jsonObj["s"]["d"]["s"] = {}
    jsonObj["s"]["d"]["s"]["name"] = data[9]
    jsonObj["s"]["d"]["d"] = {}
    jsonObj["s"]["d"]["d"]["name"] = data[10]
    jsonObj["s"]["d"]["s"]["s"] = {}
    jsonObj["s"]["d"]["s"]["s"]["name"] = data[19]
    jsonObj["s"]["d"]["s"]["d"] = {}
    jsonObj["s"]["d"]["s"]["d"]["name"] = data[20]
    jsonObj["s"]["d"]["d"]["s"] = {}
    jsonObj["s"]["d"]["d"]["s"]["name"] = data[21]
    jsonObj["s"]["d"]["d"]["d"] = {}
    jsonObj["s"]["d"]["d"]["d"]["name"] = data[22]

    jsonObj["d"] = {}
    jsonObj["d"]["name"] = data[2]
    jsonObj["d"]["s"] = {}
    jsonObj["d"]["s"]["name"] = data[5]
    jsonObj["d"]["s"]["s"] = {}
    jsonObj["d"]["s"]["s"]["name"] = data[11]
    jsonObj["d"]["s"]["d"] = {}
    jsonObj["d"]["s"]["d"]["name"] = data[12]
    jsonObj["d"]["s"]["s"]["s"] = {}
    jsonObj["d"]["s"]["s"]["s"]["name"] = data[23]
    jsonObj["d"]["s"]["s"]["d"] = {}
    jsonObj["d"]["s"]["s"]["d"]["name"] = data[24]
    jsonObj["d"]["s"]["d"]["s"] = {}
    jsonObj["d"]["s"]["d"]["s"]["name"] = data[25]
    jsonObj["d"]["s"]["d"]["d"] = {}
    jsonObj["d"]["s"]["d"]["d"]["name"] = data[26]
    jsonObj["d"]["d"] = {}
    jsonObj["d"]["d"]["name"] = data[6]
    jsonObj["d"]["d"]["s"] = {}
    jsonObj["d"]["d"]["s"]["name"] = data[13]
    jsonObj["d"]["d"]["d"] = {}
    jsonObj["d"]["d"]["d"]["name"] = data[14]
    jsonObj["d"]["d"]["s"]["s"] = {}
    jsonObj["d"]["d"]["s"]["s"]["name"] = data[27]
    jsonObj["d"]["d"]["s"]["d"] = {}
    jsonObj["d"]["d"]["s"]["d"]["name"] = data[28]
    jsonObj["d"]["d"]["d"]["s"] = {}
    jsonObj["d"]["d"]["d"]["s"]["name"] = data[29]
    jsonObj["d"]["d"]["d"]["d"] = {}
    jsonObj["d"]["d"]["d"]["d"]["name"] = data[30]

    return jsonObj

def showMessageBox(message, msg_type):
    msg = QMessageBox()
    if msg_type == MSG_ERROR:
        msg.setIcon(QMessageBox.Critical)
    elif msg_type == MSG_WARNING:
        msg.setIcon(QMessageBox.Warning)
    elif msg_type == MSG_SUCCESS:
        msg.setIcon(QMessageBox.Information)
    msg.setWindowTitle("Message")
    msg.setText(message)
    msg.setStandardButtons(QMessageBox.Ok)
    msg.exec_()