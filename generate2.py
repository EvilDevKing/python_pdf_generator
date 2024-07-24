from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select, WebDriverWait
from fpdf import FPDF, XPos, YPos, Align, svg
from constants import *
import time, math, re

cur_dir = getProjectPath()

class CustomPDF(FPDF):
    def __init__(self, orientation, unit, format):
        super().__init__(orientation, unit, format)

    def header(self):
        if self.page_no() != 1:
            # Set up a logo
            self.image('assets/images/logo_header.png', 50, 20, 60)

            # Set up a heading label
            self.set_font('Times', '', 15)
            self.set_text_color(128, 128, 128)
            self.cell(750)
            self.cell(0, 30, 'Stallion Suggestions Report', new_x=XPos.RIGHT, new_y=YPos.TOP)

            # # Line break
            self.ln(20)

def create_pdf(wsheetId=None, wsheetName=None, msheetId=None, genType=None):
    ##########################################################
    #################                    #####################
    ################# Preparing PDF Data #####################
    #################                    #####################
    ##########################################################
    url = "https://beta.allbreedpedigree.com/login"
    browser = getChromeDriver()
    browser.get(url)
    WebDriverWait(browser, 10).until(lambda browser: browser.execute_script('return document.readyState') == 'complete')
    WebDriverWait(browser, 10).until(ec.element_to_be_clickable((By.XPATH, "//button[@class='btn-close']"))).click()
    time.sleep(1)
    email_elem = browser.find_element(By.CSS_SELECTOR, "input#id-field-email")
    email_elem.click()
    email_elem.send_keys("brittany.holy@gmail.com")

    pass_elem = browser.find_element(By.CSS_SELECTOR, "input#id-field-password")
    pass_elem.click()
    pass_elem.send_keys("7f2uwvm5e4sD5PH")

    form_elem = browser.find_element(By.CSS_SELECTOR, "form#login-form")
    form_elem.submit()
    
    WebDriverWait(browser, 10).until(ec.element_to_be_clickable((By.XPATH, "//div[@id='header-search-input-helper']"))).click()
    input_elem = WebDriverWait(browser, 10).until(ec.element_to_be_clickable((By.XPATH, "//input[@id='header-search-input']")))
    input_elem.send_keys(Keys.CONTROL + "a")
    input_elem.send_keys(Keys.DELETE)
    input_elem.send_keys(wsheetName, Keys.ENTER)

    pedigree_dict = dict()
    try:
        table = browser.find_element(By.CSS_SELECTOR, "table.pedigree-table tbody")
        pedigree_dict["sex"] = browser.find_element(By.CSS_SELECTOR, "span#pedigree-animal-info span[title='Sex']").text
        pedigree_dict["birth"] = browser.find_element(By.CSS_SELECTOR, "span#pedigree-animal-info span[title='Date of Birth']").text
        pedigree_dict["pedigree"] = getPedigreeDataFromTable(table)
    except:
        try:
            table = browser.find_element(By.CSS_SELECTOR, "div.layout-table-wrapper table tbody")
            tds = table.find_elements(By.CSS_SELECTOR, "td:nth-child(1)")
            txt_vals = []
            links = []
            for td in tds:
                txt_vals.append(td.text.upper())
                links.append(td.find_element(By.CSS_SELECTOR, "a").get_attribute("href"))
            indexes = [i for i, x in enumerate(txt_vals) if x.lower() == wsheetName.lower()]
            if len(indexes) == 1:
                browser.get(links[0])
                WebDriverWait(browser, 10).until(lambda browser: browser.execute_script('return document.readyState') == 'complete')
                table = browser.find_element(By.CSS_SELECTOR, "table.pedigree-table tbody")
                pedigree_dict["sex"] = browser.find_element(By.CSS_SELECTOR, "span#pedigree-animal-info span[title='Sex']").text
                pedigree_dict["birth"] = browser.find_element(By.CSS_SELECTOR, "span#pedigree-animal-info span[title='Date of Birth']").text
                pedigree_dict["pedigree"] = getPedigreeDataFromTable(table)
            else:
                try:
                    select = Select(browser.find_element(By.XPATH, "//select[@id='filter-match']"))
                    select.select_by_value("exact")
                    WebDriverWait(browser, 10).until(lambda browser: browser.execute_script('return document.readyState') == 'complete')
                    table = browser.find_element(By.CSS_SELECTOR, "table.pedigree-table tbody")
                    pedigree_dict["sex"] = browser.find_element(By.CSS_SELECTOR, "span#pedigree-animal-info span[title='Sex']").text
                    pedigree_dict["birth"] = browser.find_element(By.CSS_SELECTOR, "span#pedigree-animal-info span[title='Date of Birth']").text
                    pedigree_dict["pedigree"] = getPedigreeDataFromTable(table)
                except:
                    browser.quit()
                    return {"status": MSG_ERROR, "msg": "Not found your horse pedigree data from allbreedpedigree.com."}
        except:
            browser.quit()
            return {"status": MSG_ERROR, "msg": "Not found your horse pedigree data from allbreedpedigree.com."}
    
    linebred_link = browser.find_element(By.CSS_SELECTOR, "div#report-menu li:nth-child(4) a").get_attribute("href")
    browser.get(linebred_link)
    gen_select = Select(browser.find_element(By.CSS_SELECTOR, "select[name='gens']"))
    gen_select.select_by_visible_text("5")
    time.sleep(1)
    try:
        coi_val = browser.find_element(By.CSS_SELECTOR, "blockquote ul li:nth-child(1) span.text-success strong").text
        coi_val = get2DigitsStringValue(float(coi_val.replace("%",""))) + "%"
    except:
        coi_val = "0.00%"

    sire_name = pedigree_dict["pedigree"][5]
    damssire_name = pedigree_dict["pedigree"][13]
    damssire2_name = pedigree_dict["pedigree"][17]
    damssire3_name = pedigree_dict["pedigree"][21]

    worksheet = getGoogleSheetService().spreadsheets()
    base_data = worksheet.values().get(spreadsheetId=wsheetId, range=f"{wsheetName}!B4:C").execute().get('values')
    base_data2 = worksheet.values().get(spreadsheetId=msheetId, range="1d crosses!B2:C").execute().get('values')
    tier1_basedata = []
    tier2_basedata = []
    tier3_basedata = []
    tier4_basedata = []
    
    sire_predicts = worksheet.values().get(spreadsheetId=msheetId, range="Prediction rating v2!BA2:BE").execute().get('values')
    sire_unique_predicts = worksheet.values().get(spreadsheetId=msheetId, range="Prediction rating v2!AN2:AO").execute().get('values')
    damssire_predicts = worksheet.values().get(spreadsheetId=msheetId, range="Prediction rating v2!BL2:BP").execute().get('values')
    damssire_unique_predicts = worksheet.values().get(spreadsheetId=msheetId, range="Prediction rating v2!AQ2:AR").execute().get('values')
    damssire2_predicts = worksheet.values().get(spreadsheetId=msheetId, range="Prediction rating v2!BV2:BZ").execute().get('values')
    damssire2_unique_predicts = worksheet.values().get(spreadsheetId=msheetId, range="Prediction rating v2!AT2:AU").execute().get('values')
    damssire3_predicts = worksheet.values().get(spreadsheetId=msheetId, range="Prediction rating v2!CF2:CJ").execute().get('values')
    damssire3_unique_predicts = worksheet.values().get(spreadsheetId=msheetId, range="Prediction rating v2!AW2:AX").execute().get('values')
    
    sire_pred = [x for x in sire_predicts if str(x[0]).lower() == sire_name.replace("*","").lower()]
    sire_unique_pred = [x for x in sire_unique_predicts if str(x[0]).lower() == sire_name.replace("*","").lower()]
    damssire_pred = [x for x in damssire_predicts if str(x[0]).lower() == damssire_name.replace("*","").lower()]
    damssire_unique_pred = [x for x in damssire_unique_predicts if str(x[0]).lower() == damssire_name.replace("*","").lower()]
    damssire2_pred = [x for x in damssire2_predicts if str(x[0]).lower() == damssire2_name.replace("*","").lower()]
    damssire2_unique_pred = [x for x in damssire2_unique_predicts if str(x[0]).lower() == damssire2_name.replace("*","").lower()]
    damssire3_pred = [x for x in damssire3_predicts if str(x[0]).lower() == damssire3_name.replace("*","").lower()]
    damssire3_unique_pred = [x for x in damssire3_unique_predicts if str(x[0]).lower() == damssire3_name.replace("*","").lower()]

    if len(sire_pred) == 0:
        oned_sire = 0
        v_sire = 0
        g_sire = "B-"
    else:
        oned_sire = sire_pred[0][1]
        v_sire = sire_pred[0][3]
        g_sire = sire_pred[0][4]
        
    if len(sire_unique_pred) == 0:
        unique_sire = 0
    else:
        unique_sire = sire_unique_pred[0][1]
        
    if len(damssire_pred) == 0:
        oned_damssire = 0
        v_damssire = 0
        g_damssire = "B-"
    else:
        oned_damssire = damssire_pred[0][1]
        v_damssire = damssire_pred[0][3]
        g_damssire = damssire_pred[0][4]
        
    if len(damssire_unique_pred) == 0:
        unique_damssire = 0
    else:
        unique_damssire = damssire_unique_pred[0][1]

    if len(damssire2_pred) == 0:
        oned_damssire2 = 0
        v_damssire2 = 0
        g_damssire2 = "B-"
    else:
        oned_damssire2 = damssire2_pred[0][1]
        v_damssire2 = damssire2_pred[0][3]
        g_damssire2 = damssire2_pred[0][4]
        
    if len(damssire2_unique_pred) == 0:
        unique_damssire2 = 0
    else:
        unique_damssire2 = damssire2_unique_pred[0][1]

    if len(damssire3_pred) == 0:
        oned_damssire3 = 0
        v_damssire3 = 0
        g_damssire3 = "B-"
    else:
        oned_damssire3 = damssire3_pred[0][1]
        v_damssire3 = damssire3_pred[0][3]
        g_damssire3 = damssire3_pred[0][4]
        
    if len(damssire3_unique_pred) == 0:
        unique_damssire3 = 0
    else:
        unique_damssire3 = damssire3_unique_pred[0][1]

    grade_info = getLetterGradeBy(g_sire, g_damssire, g_damssire2, g_damssire3)
    letter_grade = grade_info["letter"]
    grade_color = grade_info["color_info"]
    v_sum = get2DigitsStringValue(float(v_sire) + float(v_damssire) + float(v_damssire2) + float(v_damssire3))

    pivot_data = worksheet.values().get(spreadsheetId=wsheetId, range=f"{wsheetName}!U4:AD").execute().get('values')
    tier_suggestions = dict()
    tier_label = ""
    for pd in pivot_data:
        if len(pd) == 0:
            break
        if pd[0].strip() != "":
            if tier_label != pd[0].strip():
                tier_label = pd[0]
            tier_suggestions[tier_label] = []
        tier_suggestions[tier_label].append([pd[1], pd[2], pd[6], pd[7], pd[8], pd[9]])

    tier1_sugs = []
    tier2_sugs = []
    tier3_sugs = []
    tier4_sugs = []
    if "tier 1" in tier_suggestions.keys():
        tier1_sugs = tier_suggestions["tier 1"]
    if "tier 2" in tier_suggestions.keys():
        temp_tier2_sugs = tier_suggestions["tier 2"]
        for v in temp_tier2_sugs:
            if v[2].strip() == "%" and v[3].strip() == "" and v[4].strip() == "":
                tier2_sugs.append([v[0], v[1], "N/A", "N/A", "N/A", v[5]])
            else:
                tier2_sugs.append(v)
    if "tier 3" in tier_suggestions.keys():
        tier3_sugs = tier_suggestions["tier 3"]
    if "tier 4" in tier_suggestions.keys():
        tier4_sugs = tier_suggestions["tier 4"]
        
    for bd in base_data:
        if len(bd) == 0: break
        for sug in tier1_sugs:
            if bd[1].lower() == sug[1].lower():
                tier1_basedata.append(bd)
        for sug in tier3_sugs:
            if bd[1].lower() == sug[1].lower():
                tier3_basedata.append(bd)
        for sug in tier3_sugs:
            if bd[1].lower() == sug[1].lower():
                tier4_basedata.append(bd)
                
    for bd in base_data2:
        if len(bd) < 2: continue
        for sug in tier2_sugs:
            if bd[1].lower() == sug[1].lower():
                tier2_basedata.append(bd)
    
    browser.quit()
    
    ############################################################
    #################                        ###################
    ################# PDF Generation Process ###################
    #################                        ###################
    ############################################################
    lmargin = 20
    pdf = CustomPDF(orientation='L', unit='pt', format=(600, 1000))

    ################# page 1 #################
    pdf.add_page()
    pdf.image('assets/images/cover.png', 0, 0, 1000, 600)
    pdf.image('assets/images/logo_big.png', 350, 10, 190)
    pdf.ln(80)
    pdf.set_font('Times', '', 25)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(350)
    pdf.multi_cell(w=150, h=30, text="Stallion Suggestions", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.ln(280)
    pdf.set_font_size(25)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(380)
    pdf.multi_cell(w=0, h=30, text="Analysis of:", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_font_size(30)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(340)
    pdf.multi_cell(w=200, h=30, text=wsheetName, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    ################# page 2 #################
    pdf.add_page()
    pdf.ln(50)
    pdf.set_font('Times', 'B', 15)
    pdf.set_text_color(0, 50, 120)
    pdf.cell(lmargin)
    pdf.cell(w=0, h=10, text="What are statistics?", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_font_size(13)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(lmargin)
    pdf.write_html("<p line-height='1.3'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Statistical analysis is different from scientific analysis. With the scientific method, you hypothesize an outcome and test the explanation through controlled experimentation that is verified through replications. In statistics, you find a correlation or commonality within a given sample of data that takes into account the varying degrees of importance of the variables in that data set. Statistical analysis works best with larger datasets because the accuracy of the results increase with the number of observations. We have a database of over 10,000 individual horses and growing. It's important to look at the same data in multiple different ways to further help simplify those larger datasets and to identify outliers.</p>")

    pdf.ln()

    pdf.set_font('Times', 'BI', 15)
    pdf.set_text_color(0, 50, 120)
    pdf.cell(lmargin)
    pdf.cell(w=0, h=10, text="The Equi-Source Score and Rating", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_font_size(13)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(lmargin)
    pdf.write_html("<p line-height='1.3'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;The <i>Equi-Source Score</i> and <i>Rating</i> consists of a weighted algorithm of four different independent coefficients: </p>")
    pdf.write_html("<p line-height='0.3'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Sire</p>")
    pdf.write_html("<p line-height='0.3'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Dam's Sire</p>")
    pdf.write_html("<p line-height='0.3'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;second Dam's Sire</p>")
    pdf.write_html("<p line-height='0.3'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;third Dam's Sire</p>")
    pdf.set_fill_color(0, 0, 0)
    pdf.circle(65, 275, 3, style="FD")
    pdf.circle(65, 292, 3, style="FD")
    pdf.circle(65, 309, 3, style="FD")
    pdf.circle(65, 326, 3, style="FD")
    
    pdf.ln()
    pdf.cell(lmargin)
    pdf.write_html("<p line-height='1.3'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Each coefficient is calculated separately, so the same stallion could have four different scores depending on where he appears in the  progeny's pedigree. This proprietary algorithm measures four different variables that impact a stallion's success:</p>")
    pdf.write_html("<p line-height='0.3'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Total performers</p>")
    pdf.write_html("<p line-height='0.3'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbspTop performers</p>")
    pdf.write_html("<p line-height='0.3'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Number of unique top performers</p>")
    pdf.write_html("<p line-height='0.3'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Jockey experience level</p>")
    pdf.circle(65, 395, 3, style="FD")
    pdf.circle(65, 412, 3, style="FD")
    pdf.circle(65, 429, 3, style="FD")
    pdf.circle(65, 446, 3, style="FD")
    
    pdf.ln()
    pdf.cell(lmargin)
    pdf.write_html("<p line-height='1.3'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;The \"variable\" score is then translated to a letter grade rating (A+ to B-), known as the \"EquiSource Score\" using four different median splits and the average of each independent coefficient (Sire, Dam's Sire, second Dam's Sire, third Dam's Sire). The two-step process in calculating the letter grade is essential in order to eliminate the bias that naturally occurs with small sample sizes. When we compile these suggestions, we adjust the pedigree to suit the future progeny which predicts the success of a proposed breeding.</p>")

    ################# page 3 #################
    pdf.add_page()
    pdf.ln(50)
    
    pdf.set_font('Times', 'B', 15)
    pdf.set_text_color(0, 50, 120)
    pdf.cell(lmargin)
    pdf.cell(w=0, h=10, text="Tiering System", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    
    pdf.set_font_size(13)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(lmargin)
    pdf.write_html("<p line-height='1.3'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbspOur predictive analysis categorizes stallion suggestions into four tiers based on the four independent coefficients within the pedigree, as it relates to heritability aiming to optimize breeding outcomes.</p>")
    pdf.cell(lmargin+10)
    pdf.write_html("<p line-height='1'>Tier 1 represents the most favorable options, from more recent generations, showcasing immediate genetic influence.</p>")
    pdf.write_html("<p line-height='1'>Tier 2, the <i><b>Stallion Alternative</b></i> section, of the report lists the stallions with the same or similar breeding to Tier 1 stallion suggestions and may include stallions with no high-performing progeny and/or Junior stallions.</p>")
    pdf.write_html("<p line-height='1'>Tiers 3 and 4 represent a less direct, yet valuable genetic lineage, providing a broader base for understanding hereditary traits and their manifestations in performance.</p>")
    pdf.circle(65, 173, 3, style="FD")
    pdf.circle(65, 200, 3, style="FD")
    pdf.circle(65, 239, 3, style="FD")
    
    pdf.cell(-lmargin-10)
    pdf.write_html("<p line-height='1.3'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;This tiered approach enables a comprehensive assessment of heritability factors across generations, aiding in the identification of optimal breeding strategies.</p>")
    
    pdf.ln()
    
    pdf.set_font('Times', 'B', 15)
    pdf.set_text_color(0, 50, 120)
    pdf.cell(lmargin)
    pdf.cell(w=0, h=10, text="Using your Report", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(lmargin)
    pdf.set_font_size(13)
    pdf.set_text_color(0, 0, 0)
    pdf.write_html("<p line-height='1.3'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;The data we compile is obtained from the American Quarter Horse Association and online published results from individual producers. Some horse pedigrees are obtained directly from owners or riders.</p>")
    pdf.write_html("<p line-height='1.3'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;The first portion of your report shows the pedigree of the horse being analyzed with the individual algorithm score of each variable highlighted by colored boxes.</p>")
    pdf.write_html("<p line-height='1.3'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;The main section shows the top suggestions in each tier (1-4) sorted by the Equi-Source Score. The stallion suggestions in each tier are obtained using the independent variables previously discussed in your horse's unique pedigree. Each box at the top of the page shows a summary of the cumulative stallion data to further support the provided suggestions in each tier.</p>")

    ################# page 4 #################
    pdf.add_page()
    pdf.ln(50)

    pdf.set_font('Times', 'B', 15)
    pdf.set_text_color(0, 50, 120)
    pdf.cell(lmargin)
    pdf.cell(w=0, h=15, text="Interpreting the data", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    
    pdf.set_font_size(13)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(2*lmargin+10)
    pdf.write_html("<p line-height='1.3'>A stallion with many different progeny that place in the top of a class, many different times, will have a higher EquiSource Score than a stallion that only has a few different progeny place in the top of a class many different times.  This is also the same for the Dam's Sire, second Dam's Sire, and third Dam's Sire. For Example:</p>")
    pdf.circle(65, 133, 3, style="FD")
    
    pdf.cell(-lmargin-10)
    pdf.set_text_color(0, 0, 0)
    pdf.write_html("<p line-height='0.3'>Stallion A:</p>")
    pdf.write_html("<p line-height='0.3'>20 total progeny across multiple events</p>")
    pdf.write_html("<p line-height='0.3'>10 total progeny place in the top of those events (1D or top 25%)</p>")
    pdf.write_html("<p line-height='0.3'>6 of those top 10 progeny are different horses, 4 are the same exact horse</p>")
    pdf.write_html("<p line-height='0.3'>4 of those top 10 progeny are ridden by professionals and the other 6 are ridden by amateurs</p>")
    pdf.write_html("<p line-height='0.3'>Resulting EquiSource Score variable: 11.0 or A+</p>")
    pdf.ln(10)
    pdf.cell(lmargin)
    pdf.write_html("<p line-height='0.3'>Stallion B:</p>")
    pdf.write_html("<p line-height='0.3'>20 total progeny across multiple events</p>")
    pdf.write_html("<p line-height='0.3'>12 total progeny place in the top of those events (1D or top 25%)</p>")
    pdf.write_html("<p line-height='0.3'>4 of those top 12 progeny are different horses, 8 are the same exact horse</p>")
    pdf.write_html("<p line-height='0.3'>10 of those top 12 progeny are ridden by professionals and the other 2 are ridden by amateurs</p>")
    pdf.write_html("<p line-height='0.3'>Resulting EquiSource Score variable: 9.4 or A</p>")
    pdf.ln(10)
    pdf.cell(lmargin)
    pdf.write_html("<p line-height='1.2'>To translate the variable score to a letter grade, the top 5% in any coefficient is rated A+. If there are a total of 400 stallions in the Sire index coefficient, then the top 20 of those are A+.</p>")
    pdf.write_html("<p line-height='1.2'>Then we take the top half median split of stallions ranked number 21-400 for the A and A- rating and the bottom half is the B and B-.</p>")
    pdf.write_html("<p line-height='1.2'>Finally, we complete two more median splits to determine each specific letter grade. The top quarter is A, the second quarter is A-, the third quarter is B and the fourth quarter is B-.</p>")
    
    ################# page 5 (Pedigree Table) #################
    pdf.add_page()
    pdf.set_line_width(2)
    pdf.set_fill_color(r=255, g=255, b=255)
    pdf.rect(x=50, y=100, w=240, h=70, style="D")
    pdf.rect(x=390, y=100, w=90, h=70, style="D")
    pdf.image('assets/images/tempbar.png', 650, 120, 200, 60)
    pdf.image('assets/images/teardrop.png', 800, 75, 25)
    
    pdf.ln()
    
    pdf.set_font('Times', 'B', 6)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(773)
    pdf.cell(w=22, h=20, text=v_sum, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
    
    pdf.ln()
    
    pdf.set_font('Times', '', 20)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(lmargin+10)
    pdf.cell(w=240, h=40, text=wsheetName, new_x=XPos.LMARGIN, new_y=YPos.TOP)

    pdf.set_font('Times', '', 45)
    pdf.set_text_color(grade_color[0], grade_color[1], grade_color[2])
    pdf.cell(360)
    pdf.cell(w=90, h=50, text=letter_grade, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    
    pdf.set_font('Times', '', 13)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(lmargin+10)
    pdf.cell(w=280, h=10, text=f"{pedigree_dict['birth']} {pedigree_dict['sex']}", new_x=XPos.LMARGIN, new_y=YPos.TOP)

    pdf.set_font('Times', '', 9)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(360)
    pdf.cell(w=90, h=25, text=f"VARIANT = {v_sum}", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    
    pdf.set_line_width(0.5)
    pdf.ln(92)
    
    # MMM
    anc_mmm = pedigree_dict["pedigree"][0]
    pdf.cell(240)
    pdf.cell(w=100, h=0, text=anc_mmm, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
    pdf.set_draw_color(0, 0, 0)
    pdf.rect(x=270, y=254, w=100, h=22, style="D")
    # MMMM
    anc_mmmm = pedigree_dict["pedigree"][1]
    pdf.cell(350)
    pdf.cell(w=100, h=0, text=anc_mmmm, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
    pdf.set_draw_color(0, 0, 0)
    pdf.rect(x=380, y=254, w=100, h=22, style="D")
    pdf.ln(16)
    # MM
    anc_mm = pedigree_dict["pedigree"][2]
    pdf.cell(130)
    pdf.cell(w=100, h=0, text=anc_mm, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
    pdf.set_draw_color(0, 0, 0)
    pdf.rect(x=160, y=270, w=100, h=22, style="D")
    pdf.ln(15)
    # MMF
    anc_mmf = pedigree_dict["pedigree"][3]
    pdf.cell(240)
    pdf.cell(w=100, h=0, text=anc_mmf, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
    pdf.set_draw_color(0, 0, 0)
    pdf.rect(x=270, y=285, w=100, h=22, style="D")
    # MMFM
    anc_mmfm = pedigree_dict["pedigree"][4]
    pdf.cell(350)
    pdf.cell(w=100, h=0, text=anc_mmfm, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
    pdf.set_draw_color(0, 0, 0)
    pdf.rect(x=380, y=285, w=100, h=22, style="D")
    pdf.ln(15)
    # M
    anc_m = pedigree_dict["pedigree"][5]
    pdf.cell(20)
    pdf.cell(w=100, h=0, text=anc_m, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
    pdf.set_draw_color(155, 203, 244)
    pdf.rect(x=50, y=300, w=100, h=22, style="D")
    pdf.ln(15)
    # MFM
    anc_mfm = pedigree_dict["pedigree"][6]
    pdf.cell(240)
    pdf.cell(w=100, h=0, text=anc_mfm, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
    pdf.set_draw_color(0, 0, 0)
    pdf.rect(x=270, y=315, w=100, h=22, style="D")
    # MFMM
    anc_mfmm = pedigree_dict["pedigree"][7]
    pdf.cell(350)
    pdf.cell(w=100, h=0, text=anc_mfmm, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
    pdf.set_draw_color(0, 0, 0)
    pdf.rect(x=380, y=315, w=100, h=22, style="D")
    pdf.ln(15)
    # MF
    anc_mf = pedigree_dict["pedigree"][8]
    pdf.cell(130)
    pdf.cell(w=100, h=0, text=anc_mf, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
    pdf.set_draw_color(0, 0, 0)
    pdf.rect(x=160, y=330, w=100, h=22, style="D")
    pdf.ln(15)
    # MFF
    anc_mff = pedigree_dict["pedigree"][9]
    pdf.cell(240)
    pdf.cell(w=100, h=0, text=anc_mff, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
    pdf.set_draw_color(0, 0, 0)
    pdf.rect(x=270, y=345, w=100, h=22, style="D")
    # MFFM
    anc_mffm = pedigree_dict["pedigree"][10]
    pdf.cell(350)
    pdf.cell(w=100, h=0, text=anc_mffm, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_draw_color(0, 0, 0)
    pdf.rect(x=380, y=345, w=100, h=22, style="D")
    pdf.ln(30)
    # FMM
    anc_fmm = pedigree_dict["pedigree"][11]
    pdf.cell(240)
    pdf.cell(w=100, h=0, text=anc_fmm, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
    pdf.set_draw_color(0, 0, 0)
    pdf.rect(x=270, y=375, w=100, h=22, style="D")
    # FMMM
    anc_fmmm = pedigree_dict["pedigree"][12]
    pdf.cell(350)
    pdf.cell(w=100, h=0, text=anc_fmmm, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
    pdf.set_draw_color(0, 0, 0)
    pdf.rect(x=380, y=375, w=100, h=22, style="D")
    pdf.ln(15)
    # FM
    anc_fm = pedigree_dict["pedigree"][13]
    pdf.cell(130)
    pdf.cell(w=100, h=0, text=anc_fm, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
    pdf.set_draw_color(56, 224, 225)
    pdf.rect(x=160, y=390, w=100, h=22, style="D")
    pdf.ln(15)
    # FMF
    anc_fmf = pedigree_dict["pedigree"][14]
    pdf.cell(240)
    pdf.cell(w=100, h=0, text=anc_fmf, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
    pdf.set_draw_color(0, 0, 0)
    pdf.rect(x=270, y=405, w=100, h=22, style="D")
    # FMFM
    anc_fmfm = pedigree_dict["pedigree"][15]
    pdf.cell(350)
    pdf.cell(w=100, h=0, text=anc_fmfm, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
    pdf.set_draw_color(0, 0, 0)
    pdf.rect(x=380, y=405, w=100, h=22, style="D")
    pdf.ln(15)
    # F
    anc_f = pedigree_dict["pedigree"][16]
    pdf.cell(20)
    pdf.cell(w=100, h=0, text=anc_f, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
    pdf.set_draw_color(0, 0, 0)
    pdf.rect(x=50, y=420, w=100, h=22, style="D")
    pdf.ln(15)
    # FFM
    anc_ffm = pedigree_dict["pedigree"][17]
    pdf.cell(240)
    pdf.cell(w=100, h=0, text=anc_ffm, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
    pdf.set_draw_color(254, 32, 253)
    pdf.rect(x=270, y=435, w=100, h=22, style="D")
    # FFMM
    anc_ffmm = pedigree_dict["pedigree"][18]
    pdf.cell(350)
    pdf.cell(w=100, h=0, text=anc_ffmm, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
    pdf.set_draw_color(0, 0, 0)
    pdf.rect(x=380, y=435, w=100, h=22, style="D")
    pdf.ln(15)
    # FF
    anc_ff = pedigree_dict["pedigree"][19]
    pdf.cell(130)
    pdf.cell(w=100, h=0, text=anc_ff, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
    pdf.set_draw_color(0, 0, 0)
    pdf.rect(x=160, y=450, w=100, h=22, style="D")
    pdf.ln(15)
    # FFF
    anc_fff = pedigree_dict["pedigree"][20]
    pdf.cell(240)
    pdf.cell(w=100, h=0, text=anc_fff, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
    pdf.set_draw_color(0, 0, 0)
    pdf.rect(x=270, y=465, w=100, h=22, style="D")
    # FFFM
    anc_fffm = pedigree_dict["pedigree"][21]
    pdf.cell(350)
    pdf.cell(w=100, h=0, text=anc_fffm, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
    pdf.set_draw_color(250, 47, 190)
    pdf.rect(x=380, y=465, w=100, h=22, style="D")
    
    ### progeny tables
    pdf.set_xy(5, 255)
    pdf.set_left_margin(1050)
    pdf.set_draw_color(155, 203, 244)
    pdf.set_fill_color(255, 255, 255) # Back to white background
    TABLE_DATA = (
        ("Sire", "Index", "Top Progeny", "Unique Top Progeny"),
        (sire_name, v_sire, oned_sire, unique_sire)
    )

    with pdf.table(text_align=Align.C, col_widths=100, line_height=5) as table:
        for data_row in TABLE_DATA:
            row = table.row()
            for datum in data_row:
                row.cell(datum, padding=(8, 5, 8, 5))
    
    pdf.ln(20)
    pdf.set_draw_color(56, 224, 225)
    pdf.set_fill_color(255, 255, 255) # Back to white background
    TABLE_DATA = (
        ("Dam's Sire", "Index", "Top Progeny", "Unique Top Progeny"),
        (damssire_name, v_damssire, oned_damssire, unique_damssire)
    )

    with pdf.table(text_align=Align.C, col_widths=100, line_height=5) as table:
        for data_row in TABLE_DATA:
            row = table.row()
            for datum in data_row:
                row.cell(datum, padding=(8, 5, 8, 5))
                
    pdf.ln(20)
    pdf.set_draw_color(254, 32, 253)
    pdf.set_fill_color(255, 255, 255) # Back to white background
    TABLE_DATA = (
        ("2nd Dam's Sire", "Index", "Top Progeny", "Unique Top Progeny"),
        (damssire2_name, v_damssire2, oned_damssire2, unique_damssire2)
    )

    with pdf.table(text_align=Align.C, col_widths=100, line_height=5) as table:
        for data_row in TABLE_DATA:
            row = table.row()
            for datum in data_row:
                row.cell(datum, padding=(8, 5, 8, 5))
                
    pdf.ln(20)
    pdf.set_draw_color(250, 47, 190)
    pdf.set_fill_color(255, 255, 255) # Back to white background
    TABLE_DATA = (
        ("3rd Dam's Sire", "Index", "Top Progeny", "Unique Top Progeny"),
        (damssire3_name, v_damssire3, oned_damssire3, unique_damssire3)
    )

    with pdf.table(text_align=Align.C, col_widths=100, line_height=5) as table:
        for data_row in TABLE_DATA:
            row = table.row()
            for datum in data_row:
                row.cell(datum, padding=(8, 5, 8, 5))
                
    pdf.set_left_margin(0)
                
    ################# page 6 (Tier 1 Suggestions) #################
    if len(tier1_sugs) == 0:
        pdf.add_page()
        pdf.set_line_width(2)
        pdf.set_draw_color(0)
        pdf.set_fill_color(r=255, g=255, b=255)
        pdf.rect(x=50, y=100, w=240, h=60, style="D")
        pdf.rect(x=360, y=100, w=90, h=60, style="D")
        pdf.rect(x=520, y=100, w=90, h=60, style="D")
        pdf.rect(x=680, y=100, w=90, h=60, style="D")
        pdf.rect(x=840, y=100, w=90, h=60, style="D")

        pdf.ln(65)
        pdf.set_font('Times', '', 22)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(lmargin+65)
        pdf.cell(w=0, h=30, text="Tier 1 Suggestions", new_x=XPos.LMARGIN, new_y=YPos.TOP)
        
        pdf.cell(360)
        pdf.set_font_size(50)
        pdf.cell(w=90, h=25, text="0", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        
        pdf.cell(520)
        pdf.set_font_size(50)
        pdf.cell(w=90, h=25, text="0", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        
        pdf.cell(680)
        pdf.set_font_size(50)
        pdf.cell(w=90, h=25, text="0", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        
        pdf.cell(840)
        pdf.set_font_size(50)
        pdf.cell(w=90, h=25, text="0", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        
        pdf.ln()
        pdf.set_font_size(10)
        pdf.cell(360)
        pdf.cell(w=90, h=30, text="MATCHES FOUND", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.cell(520)
        pdf.cell(w=90, h=30, text="EVENTS", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.cell(680)
        pdf.cell(w=90, h=30, text="TOP PLACINGS", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.cell(840)
        pdf.cell(w=90, h=30, text="PROGENY", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        
        pdf.ln(100)

        pdf.set_font('Times', '', 15)
        pdf.cell(lmargin)
        pdf.cell(w=0, h=0, text="NO TIER 1 STALLION SUGGESTIONS FOUND.", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        
    else:
        sorted_tier1_sugs = sortByOtherTiers(tier1_sugs, genType)
        sorted_tier1_sugs = [x[1:] for x in sorted_tier1_sugs]
        for i in range(math.ceil(len(sorted_tier1_sugs) / 10)):
            pdf.add_page()
            pdf.set_line_width(2)
            pdf.set_draw_color(0)
            pdf.set_fill_color(r=255, g=255, b=255)
            pdf.rect(x=50, y=100, w=240, h=60, style="D")
            pdf.rect(x=360, y=100, w=90, h=60, style="D")
            pdf.rect(x=520, y=100, w=90, h=60, style="D")
            pdf.rect(x=680, y=100, w=90, h=60, style="D")
            pdf.rect(x=840, y=100, w=90, h=60, style="D")

            pdf.ln(65)
            pdf.set_font('Times', '', 22)
            pdf.set_text_color(0, 0, 0)
            pdf.cell(lmargin+65)
            pdf.cell(w=0, h=30, text="Tier 1 Suggestions", new_x=XPos.LMARGIN, new_y=YPos.TOP)
            
            pdf.cell(360)
            pdf.set_font_size(50)
            pdf.cell(w=90, h=25, text=str(len(tier1_sugs)), align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            
            pdf.cell(520)
            pdf.set_font_size(50)
            pdf.cell(w=90, h=25, text="27", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            
            pdf.cell(680)
            pdf.set_font_size(50)
            pdf.cell(w=90, h=25, text="140", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            
            pdf.cell(840)
            pdf.set_font_size(50)
            pdf.cell(w=90, h=25, text="27", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            
            pdf.ln()
            pdf.set_font_size(10)
            pdf.cell(360)
            pdf.cell(w=90, h=30, text="MATCHES FOUND", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            pdf.cell(520)
            pdf.cell(w=90, h=30, text="EVENTS", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            pdf.cell(680)
            pdf.cell(w=90, h=30, text="TOP PLACINGS", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            pdf.cell(840)
            pdf.cell(w=90, h=30, text="PROGENY", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            
            pdf.set_line_width(0.5)
            pdf.set_xy(5, 215)
            pdf.set_left_margin(70)
            pdf.set_draw_color(0)
            pdf.set_fill_color(255, 255, 255) # Back to white background
            
            
            
            TABLE_HEADER_DATA = [
                ["Stallion", "1D Rate", "Variant", "Equi-Source Score", "Inbreeding Coefficient of foal"]
            ]
            TABLE_DATA = TABLE_HEADER_DATA + sorted_tier1_sugs[i*10:i*10+10]

            with pdf.table(text_align=Align.C, col_widths=90, line_height=10) as table:
                for data_row in TABLE_DATA:
                    row = table.row()
                    for datum in data_row:
                        row.cell(datum, padding=(5, 0, 5, 0))
                        
            pdf.set_xy(5, 215)
            pdf.set_left_margin(1070)
            TABLE_HEADER_DATA = [["Top 10 Offspring", "Sire", "Top Placings", "Earnings"]]
            TABLE_DATA = TABLE_HEADER_DATA + groupBySireAndCountHorse(tier1_basedata, genType)

            with pdf.table(text_align=Align.C, col_widths=100, line_height=10) as table:
                for data_row in TABLE_DATA:
                    row = table.row()
                    for datum in data_row:
                        row.cell(datum, padding=(5, 0, 5, 0))
                    
            pdf.set_left_margin(0)
    
    ################# page 7 (Tier 2 Suggestions) #################
    if len(tier2_sugs) == 0:
        pdf.add_page()
        pdf.set_line_width(2)
        pdf.set_draw_color(0)
        pdf.set_fill_color(r=255, g=255, b=255)
        pdf.rect(x=50, y=100, w=240, h=60, style="D")
        pdf.rect(x=360, y=100, w=90, h=60, style="D")
        pdf.rect(x=520, y=100, w=90, h=60, style="D")
        pdf.rect(x=680, y=100, w=90, h=60, style="D")
        pdf.rect(x=840, y=100, w=90, h=60, style="D")

        pdf.ln(65)
        pdf.set_font('Times', '', 22)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(lmargin+65)
        pdf.cell(w=0, h=30, text="Tier 2 Suggestions", new_x=XPos.LMARGIN, new_y=YPos.TOP)
        
        pdf.cell(360)
        pdf.set_font_size(50)
        pdf.cell(w=90, h=25, text="0", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        
        pdf.cell(520)
        pdf.set_font_size(50)
        pdf.cell(w=90, h=25, text="0", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        
        pdf.cell(680)
        pdf.set_font_size(50)
        pdf.cell(w=90, h=25, text="0", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        
        pdf.cell(840)
        pdf.set_font_size(50)
        pdf.cell(w=90, h=25, text="0", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        
        pdf.ln()
        pdf.set_font_size(10)
        pdf.cell(360)
        pdf.cell(w=90, h=30, text="MATCHES FOUND", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.cell(520)
        pdf.cell(w=90, h=30, text="EVENTS", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.cell(680)
        pdf.cell(w=90, h=30, text="TOP PLACINGS", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.cell(840)
        pdf.cell(w=90, h=30, text="PROGENY", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        
        pdf.ln(100)

        pdf.set_font('Times', '', 15)
        pdf.cell(lmargin)
        pdf.cell(w=0, h=0, text="NO TIER 2 STALLION SUGGESTIONS FOUND.", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    else:
        sorted_tier2_sugs = sortByVariant2(tier2_sugs, genType)
        sorted_tier2_sugs = [x[1:] for x in sorted_tier2_sugs]
        for i in range(math.ceil(len(sorted_tier2_sugs) / 10)):
            pdf.add_page()
            pdf.set_line_width(2)
            pdf.set_draw_color(0)
            pdf.set_fill_color(r=255, g=255, b=255)
            pdf.rect(x=50, y=100, w=240, h=60, style="D")
            pdf.rect(x=360, y=100, w=90, h=60, style="D")
            pdf.rect(x=520, y=100, w=90, h=60, style="D")
            pdf.rect(x=680, y=100, w=90, h=60, style="D")
            pdf.rect(x=840, y=100, w=90, h=60, style="D")

            pdf.ln(65)
            pdf.set_font('Times', '', 22)
            pdf.set_text_color(0, 0, 0)
            pdf.cell(lmargin+65)
            pdf.cell(w=0, h=30, text="Tier 2 Suggestions", new_x=XPos.LMARGIN, new_y=YPos.TOP)
            
            pdf.cell(360)
            pdf.set_font_size(50)
            pdf.cell(w=90, h=25, text=str(len(tier2_sugs)), align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            
            pdf.cell(520)
            pdf.set_font_size(50)
            pdf.cell(w=90, h=25, text="20", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            
            pdf.cell(680)
            pdf.set_font_size(50)
            pdf.cell(w=90, h=25, text="76", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            
            pdf.cell(840)
            pdf.set_font_size(50)
            pdf.cell(w=90, h=25, text="17", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            
            pdf.ln()
            pdf.set_font_size(10)
            pdf.cell(360)
            pdf.cell(w=90, h=30, text="MATCHES FOUND", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            pdf.cell(520)
            pdf.cell(w=90, h=30, text="EVENTS", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            pdf.cell(680)
            pdf.cell(w=90, h=30, text="TOP PLACINGS", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            pdf.cell(840)
            pdf.cell(w=90, h=30, text="PROGENY", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            
            pdf.set_line_width(0.5)
            pdf.set_xy(5, 215)
            pdf.set_left_margin(70)
            pdf.set_draw_color(0)
            pdf.set_fill_color(255, 255, 255) # Back to white background
            
            TABLE_HEADER_DATA = [
                ["Stallion", "1D Rate", "Variant", "Equi-Source Score", "Inbreeding Coefficient of foal"]
            ]
            TABLE_DATA = TABLE_HEADER_DATA + sorted_tier2_sugs[i*10:i*10+10]

            with pdf.table(text_align=Align.C, col_widths=90, line_height=10) as table:
                for data_row in TABLE_DATA:
                    row = table.row()
                    for datum in data_row:
                        row.cell(datum, padding=(5, 0, 5, 0))
                        
            pdf.set_xy(5, 215)
            pdf.set_left_margin(1070)
            TABLE_HEADER_DATA = [["Top 10 Offspring", "Sire", "Top Placings", "Earnings"]]
            TABLE_DATA = TABLE_HEADER_DATA + groupBySireAndCountHorse(tier2_basedata, genType)

            with pdf.table(text_align=Align.C, col_widths=100, line_height=10) as table:
                for data_row in TABLE_DATA:
                    row = table.row()
                    for datum in data_row:
                        row.cell(datum, padding=(5, 0, 5, 0))
                    
            pdf.set_left_margin(0)
    
    
    ################# page 8 (Tier 3 Suggestions) #################
    if len(tier3_sugs) == 0:
        pdf.add_page()
        pdf.set_line_width(2)
        pdf.set_draw_color(0)
        pdf.set_fill_color(r=255, g=255, b=255)
        pdf.rect(x=50, y=100, w=240, h=60, style="D")
        pdf.rect(x=360, y=100, w=90, h=60, style="D")
        pdf.rect(x=520, y=100, w=90, h=60, style="D")
        pdf.rect(x=680, y=100, w=90, h=60, style="D")
        pdf.rect(x=840, y=100, w=90, h=60, style="D")

        pdf.ln(65)
        pdf.set_font('Times', '', 22)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(lmargin+65)
        pdf.cell(w=0, h=30, text="Tier 3 Suggestions", new_x=XPos.LMARGIN, new_y=YPos.TOP)
        
        pdf.cell(360)
        pdf.set_font_size(50)
        pdf.cell(w=90, h=25, text="0", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        
        pdf.cell(520)
        pdf.set_font_size(50)
        pdf.cell(w=90, h=25, text="0", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        
        pdf.cell(680)
        pdf.set_font_size(50)
        pdf.cell(w=90, h=25, text="0", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        
        pdf.cell(840)
        pdf.set_font_size(50)
        pdf.cell(w=90, h=25, text="0", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        
        pdf.ln()
        pdf.set_font_size(10)
        pdf.cell(360)
        pdf.cell(w=90, h=30, text="MATCHES FOUND", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.cell(520)
        pdf.cell(w=90, h=30, text="EVENTS", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.cell(680)
        pdf.cell(w=90, h=30, text="TOP PLACINGS", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.cell(840)
        pdf.cell(w=90, h=30, text="PROGENY", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        
        pdf.ln(100)

        pdf.set_font('Times', '', 15)
        pdf.cell(lmargin)
        pdf.cell(w=0, h=0, text="NO TIER 3 STALLION SUGGESTIONS FOUND.", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    else:
        sorted_tier3_sugs = sortByVariant2(tier3_sugs, genType)
        sorted_tier3_sugs = [x[1:] for x in sorted_tier3_sugs]
        for i in range(math.ceil(len(sorted_tier3_sugs) / 10)):
            pdf.add_page()
            pdf.set_line_width(2)
            pdf.set_draw_color(0)
            pdf.set_fill_color(r=255, g=255, b=255)
            pdf.rect(x=50, y=100, w=240, h=60, style="D")
            pdf.rect(x=360, y=100, w=90, h=60, style="D")
            pdf.rect(x=520, y=100, w=90, h=60, style="D")
            pdf.rect(x=680, y=100, w=90, h=60, style="D")
            pdf.rect(x=840, y=100, w=90, h=60, style="D")

            pdf.ln(65)
            pdf.set_font('Times', '', 22)
            pdf.set_text_color(0, 0, 0)
            pdf.cell(lmargin+65)
            pdf.cell(w=0, h=30, text="Tier 3 Suggestions", new_x=XPos.LMARGIN, new_y=YPos.TOP)
            
            pdf.cell(360)
            pdf.set_font_size(50)
            pdf.cell(w=90, h=25, text=str(len(tier3_sugs)), align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            
            pdf.cell(520)
            pdf.set_font_size(50)
            pdf.cell(w=90, h=25, text="28", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            
            pdf.cell(680)
            pdf.set_font_size(50)
            pdf.cell(w=90, h=25, text="288", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            
            pdf.cell(840)
            pdf.set_font_size(50)
            pdf.cell(w=90, h=25, text="63", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            
            pdf.ln()
            pdf.set_font_size(10)
            pdf.cell(360)
            pdf.cell(w=90, h=30, text="MATCHES FOUND", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            pdf.cell(520)
            pdf.cell(w=90, h=30, text="EVENTS", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            pdf.cell(680)
            pdf.cell(w=90, h=30, text="TOP PLACINGS", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            pdf.cell(840)
            pdf.cell(w=90, h=30, text="PROGENY", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            
            pdf.set_line_width(0.5)
            pdf.set_xy(5, 215)
            pdf.set_left_margin(70)
            pdf.set_draw_color(0)
            pdf.set_fill_color(255, 255, 255) # Back to white background
            
            TABLE_HEADER_DATA = [
                ["Stallion", "1D Rate", "Variant", "Equi-Source Score", "Inbreeding Coefficient of foal"]
            ]
            TABLE_DATA = TABLE_HEADER_DATA + sorted_tier3_sugs[i*10:i*10+10]

            with pdf.table(text_align=Align.C, col_widths=90, line_height=10) as table:
                for data_row in TABLE_DATA:
                    row = table.row()
                    for datum in data_row:
                        row.cell(datum, padding=(5, 0, 5, 0))
                        
            pdf.set_xy(5, 215)
            pdf.set_left_margin(1070)
            TABLE_HEADER_DATA = [["Top 10 Offspring", "Sire", "Top Placings", "Earnings"]]
            TABLE_DATA = TABLE_HEADER_DATA + groupBySireAndCountHorse(tier3_basedata, genType)

            with pdf.table(text_align=Align.C, col_widths=100, line_height=10) as table:
                for data_row in TABLE_DATA:
                    row = table.row()
                    for datum in data_row:
                        row.cell(datum, padding=(5, 0, 5, 0))
                    
            pdf.set_left_margin(0)
    
    ################# page 9 (Tier 4 Suggestions) #################
    if len(tier4_sugs) == 0:
        pdf.add_page()
        pdf.set_line_width(2)
        pdf.set_draw_color(0)
        pdf.set_fill_color(r=255, g=255, b=255)
        pdf.rect(x=50, y=100, w=240, h=60, style="D")
        pdf.rect(x=360, y=100, w=90, h=60, style="D")
        pdf.rect(x=520, y=100, w=90, h=60, style="D")
        pdf.rect(x=680, y=100, w=90, h=60, style="D")
        pdf.rect(x=840, y=100, w=90, h=60, style="D")

        pdf.ln(65)
        pdf.set_font('Times', '', 22)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(lmargin+65)
        pdf.cell(w=0, h=30, text="Tier 4 Suggestions", new_x=XPos.LMARGIN, new_y=YPos.TOP)
        
        pdf.cell(360)
        pdf.set_font_size(50)
        pdf.cell(w=90, h=25, text="0", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        
        pdf.cell(520)
        pdf.set_font_size(50)
        pdf.cell(w=90, h=25, text="0", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        
        pdf.cell(680)
        pdf.set_font_size(50)
        pdf.cell(w=90, h=25, text="0", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        
        pdf.cell(840)
        pdf.set_font_size(50)
        pdf.cell(w=90, h=25, text="0", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        
        pdf.ln()
        pdf.set_font_size(10)
        pdf.cell(360)
        pdf.cell(w=90, h=30, text="MATCHES FOUND", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.cell(520)
        pdf.cell(w=90, h=30, text="EVENTS", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.cell(680)
        pdf.cell(w=90, h=30, text="TOP PLACINGS", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.cell(840)
        pdf.cell(w=90, h=30, text="PROGENY", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        
        pdf.ln(100)

        pdf.set_font('Times', '', 15)
        pdf.cell(lmargin)
        pdf.cell(w=0, h=0, text="NO TIER 4 STALLION SUGGESTIONS FOUND.", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    else:
        sorted_tier4_sugs = sortByVariant2(tier4_sugs, genType)
        sorted_tier4_sugs = [x[1:] for x in sorted_tier4_sugs]
        for i in range(math.ceil(len(sorted_tier4_sugs) / 10)):
            pdf.add_page()
            pdf.set_line_width(2)
            pdf.set_draw_color(0)
            pdf.set_fill_color(r=255, g=255, b=255)
            pdf.rect(x=50, y=100, w=240, h=60, style="D")
            pdf.rect(x=360, y=100, w=90, h=60, style="D")
            pdf.rect(x=520, y=100, w=90, h=60, style="D")
            pdf.rect(x=680, y=100, w=90, h=60, style="D")
            pdf.rect(x=840, y=100, w=90, h=60, style="D")

            pdf.ln(65)
            pdf.set_font('Times', '', 22)
            pdf.set_text_color(0, 0, 0)
            pdf.cell(lmargin+65)
            pdf.cell(w=0, h=30, text="Tier 4 Suggestions", new_x=XPos.LMARGIN, new_y=YPos.TOP)
            
            pdf.cell(360)
            pdf.set_font_size(50)
            pdf.cell(w=90, h=25, text=str(len(tier4_sugs)), align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            
            pdf.cell(520)
            pdf.set_font_size(50)
            pdf.cell(w=90, h=25, text="29", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            
            pdf.cell(680)
            pdf.set_font_size(50)
            pdf.cell(w=90, h=25, text="456", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            
            pdf.cell(840)
            pdf.set_font_size(50)
            pdf.cell(w=90, h=25, text="90", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            
            pdf.ln()
            pdf.set_font_size(10)
            pdf.cell(360)
            pdf.cell(w=90, h=30, text="MATCHES FOUND", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            pdf.cell(520)
            pdf.cell(w=90, h=30, text="EVENTS", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            pdf.cell(680)
            pdf.cell(w=90, h=30, text="TOP PLACINGS", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            pdf.cell(840)
            pdf.cell(w=90, h=30, text="PROGENY", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
            
            pdf.set_line_width(0.5)
            pdf.set_xy(5, 215)
            pdf.set_left_margin(70)
            pdf.set_draw_color(0)
            pdf.set_fill_color(255, 255, 255) # Back to white background
            
            TABLE_HEADER_DATA = [
                ["Stallion", "1D Rate", "Variant", "Equi-Source Score", "Inbreeding Coefficient of foal"]
            ]
            TABLE_DATA = TABLE_HEADER_DATA + sorted_tier4_sugs[i*10:i*10+10]

            with pdf.table(text_align=Align.C, col_widths=90, line_height=10) as table:
                for data_row in TABLE_DATA:
                    row = table.row()
                    for datum in data_row:
                        row.cell(datum, padding=(5, 0, 5, 0))
                        
            pdf.set_xy(5, 215)
            pdf.set_left_margin(1070)
            TABLE_HEADER_DATA = [["Top 10 Offspring", "Sire", "Top Placings", "Earnings"]]
            TABLE_DATA = TABLE_HEADER_DATA + groupBySireAndCountHorse(tier4_basedata, genType)

            with pdf.table(text_align=Align.C, col_widths=100, line_height=10) as table:
                for data_row in TABLE_DATA:
                    row = table.row()
                    for datum in data_row:
                        row.cell(datum, padding=(5, 0, 5, 0))
                    
            pdf.set_left_margin(0)

    pdf.output(f"{wsheetName}.pdf")
    return {"status": MSG_SUCCESS, "msg": "Success"}
    
# create_pdf(wsheetId="1h-tZdm0-UJnC09j8dYidTND1FCWRGDxkBMCmHzr1bYM", wsheetName="Mistys Money N Fame", msheetId="18wZ_UlyQKmhzygdb8nk8I6xAyIPvxJm3Ofh58d1NKZs", genType=1)