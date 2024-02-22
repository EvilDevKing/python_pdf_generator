from selenium.webdriver import Chrome
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from fpdf import FPDF, XPos, YPos, Align
from constants import *
import time, math, os, json

cur_dir = getProjectPath()

class CustomPDF(FPDF):
    def __init__(self, orientation, unit, format):
        super().__init__(orientation, unit, format)

    def header(self):
        if self.page_no() != 1:
            # Set up a logo
            self.image('assets/images/logo_header.png', 50, 20, 40)

            # Set up a heading label
            self.set_font('Helvetica', '', 12)
            self.set_text_color(128, 128, 128)
            self.cell(400)
            self.cell(0, 20, 'Stallion Suggestions Report', new_x=XPos.RIGHT, new_y=YPos.TOP)

            # # Line break
            self.ln(20)

    def footer(self):
        # Position cursor at 1.5 cm from bottom:
        self.set_y(-15)
        # Setting font: helvetica italic 8
        self.set_font("Helvetica", "I", 8)
        # Printing page number:
        self.cell(0, 10, f"Page {self.page_no()}/{{nb}}", align="C")

def create_pdf(wsheetId=None, sheetName=None, msheetId=None, genType=None):
    ################# Preparing PDF Data #####################
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
    input_elem.send_keys(sheetName, Keys.ENTER)

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
            indexes = [i for i, x in enumerate(txt_vals) if x.lower() == sheetName.lower()]
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
    coi_val = browser.find_element(By.CSS_SELECTOR, "blockquote ul li:nth-child(1) span.text-success strong").text
    coi_val = str(round(float(coi_val.replace("%","")), 2)) + "%"

    sire_name = pedigree_dict["pedigree"][5]
    damssire_name = pedigree_dict["pedigree"][13]
    damssire2_name = pedigree_dict["pedigree"][17]
    damssire3_name = pedigree_dict["pedigree"][21]

    worksheet = getGoogleSheetService().spreadsheets()
    sire_predicts = worksheet.values().get(spreadsheetId=msheetId, range="Prediction rating v2!BA2:BE").execute().get('values')
    damssire_predicts = worksheet.values().get(spreadsheetId=msheetId, range="Prediction rating v2!BL2:BP").execute().get('values')
    damssire2_predicts = worksheet.values().get(spreadsheetId=msheetId, range="Prediction rating v2!BV2:BZ").execute().get('values')
    damssire3_predicts = worksheet.values().get(spreadsheetId=msheetId, range="Prediction rating v2!CF2:CJ").execute().get('values')
    
    sire_pred = [x for x in sire_predicts if str(x[0]).lower() == sire_name.replace("*","").lower()]
    damssire_pred = [x for x in damssire_predicts if str(x[0]).lower() == damssire_name.replace("*","").lower()]
    damssire2_pred = [x for x in damssire2_predicts if str(x[0]).lower() == damssire2_name.replace("*","").lower()]
    damssire3_pred = [x for x in damssire3_predicts if str(x[0]).lower() == damssire3_name.replace("*","").lower()]

    if len(sire_pred) == 0:
        v_sire = 0
        g_sire = "B-"
    else:
        v_sire = sire_pred[0][3]
        g_sire = sire_pred[0][4]
        
    if len(damssire_pred) == 0:
        v_damssire = 0
        g_damssire = "B-"
    else:
        v_damssire = damssire_pred[0][3]
        g_damssire = damssire_pred[0][4]

    if len(damssire2_pred) == 0:
        v_damssire2 = 0
        g_damssire2 = "B-"
    else:
        v_damssire2 = damssire2_pred[0][3]
        g_damssire2 = damssire2_pred[0][4]

    if len(damssire3_pred) == 0:
        v_damssire3 = 0
        g_damssire3 = "B-"
    else:
        v_damssire3 = damssire3_pred[0][3]
        g_damssire3 = damssire3_pred[0][4]

    grade_info = getLetterGradeBy(g_sire, g_damssire, g_damssire2, g_damssire3)
    letter_grade = grade_info["letter"]
    grade_color = grade_info["color_info"]
    v_sum = get2DigitsStringValue(float(v_sire) + float(v_damssire) + float(v_damssire2) + float(v_damssire3))

    baby_data = worksheet.values().get(spreadsheetId=wsheetId, range=f"{sheetName}!A3:R3").execute().get('values')
    tier1_filter_label = baby_data[0][5] + ", " + baby_data[0][7]
    tier3_filter_label = baby_data[0][9]
    tier4_filter_label = baby_data[0][14] + ", " + baby_data[0][15] + ", " + baby_data[0][16] + ", " + baby_data[0][17]

    pivot_data = worksheet.values().get(spreadsheetId=wsheetId, range=f"{sheetName}!U4:AD").execute().get('values')
    if "tier" not in pivot_data[0][0]:
        browser.quit()
        return {"status": MSG_ERROR, "msg": "There is no any pivot table on sheet."}
    tier_suggestions = dict()
    tier_label = ""
    for pd in pivot_data:
        if len(pd) == 0:
            break
        if pd[0].strip() != "":
            if tier_label != pd[0].strip():
                tier_label = pd[0]
            tier_suggestions[tier_label] = []
        tier_suggestions[tier_label].append([pd[2], pd[6], pd[7], pd[8], pd[9]])

    tier1_sugs = []
    tier2_sugs = []
    tier3_sugs = []
    tier4_sugs = []
    if "tier 1" in tier_suggestions.keys():
        tier1_sugs = tier_suggestions["tier 1"]
    if "tier 2" in tier_suggestions.keys():
        tier2_sugs = tier_suggestions["tier 2"]
    if "tier 3" in tier_suggestions.keys():
        tier3_sugs = tier_suggestions["tier 3"]
    if "tier 4" in tier_suggestions.keys():
        tier4_sugs = tier_suggestions["tier 4"]

    # Ancestor part
    acol_values = worksheet.values().get(spreadsheetId=wsheetId, range=f"{sheetName}!A:A").execute().get('values')
    try:
        ind_anc = acol_values.index(["Ancestors"])+1
    except:
        return {"status": MSG_ERROR, "msg": "There is no Ancestors analysis table."}
    
    fcol_values = worksheet.values().get(spreadsheetId=wsheetId, range=f"{sheetName}!F:F").execute().get('values')
    try:
        ind_sta = fcol_values.index(["Stallions"])+1
    except:
        return {"status": MSG_ERROR, "msg": "There is no Stallions Analysis table."}

    master_stallion_data = worksheet.values().get(spreadsheetId=msheetId, range=f"Stallion master pedigree!A2:AF").execute().get('values')
    tmp_anc_top_data = worksheet.values().get(spreadsheetId=wsheetId, range=f"{sheetName}!F{ind_anc}:Y").execute().get('values')
    anc_top_data = []
    anc_pedigree_data = []
    for v in tmp_anc_top_data:
        if len(v) == 0: break
        if v[0] == "Ancestors": continue
        if v[0] == "Total":
            anc_pedigree_data = v
        else:
            filtered_sire = [sire for sire in master_stallion_data if sire[0].lower() == v[0].lower()]
            if len(filtered_sire) == 0:
                anc_top_data.append([v[0], v[1], v[17], v[19], ""])
            else:
                anc_top_data.append([v[0], v[1], v[17], v[19], filtered_sire[0][-1]])

    stallion_data = worksheet.values().get(spreadsheetId=wsheetId, range=f"{sheetName}!F{ind_sta+1}:G").execute().get('values')

    browser.quit()


    ################# PDF Generation Process ###################
    lmargin = 20
    pdf = CustomPDF(orientation='P', unit='pt', format='Letter')
    pdf.alias_nb_pages()

    ################# page 1 #################
    pdf.add_page()
    pdf.image('assets/images/bg1.png', 0, 100, 620)
    pdf.image('assets/images/bg3.png', 280, 160, 380)
    pdf.image('assets/images/bg2.png', 0, 240, 620)
    pdf.image('assets/images/bg4.png', 0, 320, 300)
    pdf.image('assets/images/logo_big.png', 150, 30, 220)
    pdf.ln(100)
    pdf.set_font('Times', '', 35)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(120)
    pdf.multi_cell(w=200, h=50, text="Stallion Suggestions", align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.ln(320)
    pdf.set_font('Times', '', 35)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(200)
    pdf.multi_cell(w=0, h=50, text="Analysis of:", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_font('Times', '', 45)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(100)
    pdf.multi_cell(w=370, h=50, text=sheetName, align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    ################# page 2 #################
    pdf.add_page()
    pdf.set_font('Times', 'B', 15)
    pdf.set_text_color(0, 50, 120)
    pdf.cell(w=0, h=50, text="Equi-Source Explanation", align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.cell(lmargin)
    pdf.cell(w=0, h=10, text="Then Vs. Now", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_font('Helvetica', '', 12)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(lmargin)
    pdf.write_html("<p line-height='1.3'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;In the past, the Equine industry has had to rely on the simple intuition of breeders to create what are known as \"magic crosses\" that were based on trial and error. Not much has changed since we started keeping records of equine pedigrees, until now. Digitized records and event results have made it possible to measure and rate the most successful horses, providing value to the analysis of the outcome which Equi-Source provides.</p>")

    pdf.ln()

    pdf.set_font('Times', 'B', 15)
    pdf.set_text_color(0, 50, 120)
    pdf.cell(lmargin)
    pdf.cell(w=0, h=15, text="What are statistics?", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_font('Helvetica', '', 12)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(lmargin)
    pdf.write_html("<p line-height='1.3'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Statistical analysis is different from scientific analysis. With the scientific method, you hypothesize an outcome and test the explanation through controlled experimentation that is verified through replications. In statistics, you find a correlation or commonality within a given sample of data that takes into account the varying degrees of importance of the variables in that data set. Statistical analysis works best with larger datasets because the accuracy of the results increase with the number of observations. We have a database of over 10,000 individual horses and growing. It's important to look at the same data in multiple different ways to further help simplify those larger datasets and to identify outliers.</p>")

    pdf.ln()

    pdf.set_font('Times', 'BI', 15)
    pdf.set_text_color(0, 50, 120)
    pdf.cell(lmargin)
    pdf.cell(w=0, h=15, text="The Equi-Source Score and Rating", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_font('Helvetica', '', 12)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(lmargin)
    pdf.write_html("<p line-height='1.3'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;The <i>Equi-Source Score</i> and <i>Rating</i> consists of a weighted algorithm of four different independent coefficients with any given horse's five generation pedigree: the Sire, the Dam's Sire, the second Dam's Sire, and the third Dam's Sire. These coefficients were determined using a partition based model through SAS technology. Each coefficient is calculated separately and compares the percentage of top performers to the total field and the number of unique performers within a given event. That score is then translated to a letter grade rating (A-D) based on the median and average scores in each event. The two-step process in calculating the letter grade is essential in order to eliminate the bias that occurs with small sample sizes. When we compile these suggestions, we adjust the pedigree to suit the future progeny so that your mare's Sire becomes the Dam's Sire and so on and so forth, which predicts the success of a proposed breeding.</p>")

    pdf.image('assets/images/logo_grades.png', x=Align.C, y=600, w=250)

    ################# page 3 #################
    pdf.add_page()
    pdf.ln()

    pdf.set_font('Helvetica', '', 12)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(lmargin)
    pdf.write_html("<p line-height='1.3'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;We also include the <i>Inbreeding Coefficient</i> based on Sewall Wright's \"F Calculation\" that is expressed as a percentage to illustrate genetic diversity or lack thereof. In 2018, a peer reviewed study conducted by Evelyn Todd of The University of Sydney (AUS) and Brandon Veilie of Swedish University of Agricultural Sciences (SWE) with data from Racing Australia and The Australian Stud Book, found that higher levels of inbreeding had a negative effect on racing performance. Horses are somewhat inbred if the Inbreeding Coefficient is 1.0, and extremely inbred if the Inbreeding Coefficient is 5.35.</p>")
    pdf.write_html("<p line-height='1.3'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;In an effort to provide as many stallion suggestions as possible, and since our data is limited by what's digitally available, the <i>Stallion Alternative</i> section of the report lists the most prolific ancestors appearing in the suggested stallion's pedigree. You can use this section to help you find similarly bred stallions that Equi-Source may not have captured due to geographic location, discipline or time standing at stud as an alternative to the suggested stallions.</p>")
    pdf.write_html("<p line-height='1.3'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Not only are these ancestors individually significant, but their position in the pedigree is as well. We already know that some sires produce better broodmares than stallions, but the most valuable ones can be used in a diverse range of positions on the pedigree which is illustrated in the \"Position Diversity\" column of that table. The whole number indicates the different number of positions that ancestors are found in the cumulative list of suggested stallions.</p>")
    pdf.write_html("<p line-height='1.3'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;The \"Position Flexibility Score\" is a weighted algorithm that compares three variables: frequency, diversity and position in the suggested stallions' pedigree. The higher the score, the more valuable the ancestor is. Theoretically, the more \"Top 10 Ancestors\" a suggested stallion has, the more valuable they are.</p>")
    pdf.write_html("<p line-height='1.3'>Reference: Todd ET, Ho SYW, Thomson PC, Ang RA, Velie BD, Hamilton NA. Founder-specific inbreeding depression affects racing performance in Thoroughbred horses. Sci Rep. 2018 Apr 18;8(1):6167. doi: 10.1038/s41598-018-24663-x. PMID: 29670190; PMCID: PMC5906619.</p>")

    pdf.ln()

    pdf.set_font('Times', 'B', 15)
    pdf.set_text_color(0, 50, 120)
    pdf.cell(lmargin)
    pdf.cell(w=0, h=15, text="Using your report", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.ln()

    pdf.set_font('Helvetica', '', 12)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(lmargin)
    pdf.write_html("<p line-height='1.3'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;The data we compile is obtained from the American Quarter Horse Association and online published results from individual producers. Some horse pedigrees are obtained directly from owners or riders.</p>")
    pdf.write_html("<p line-height='1.3'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;The first portion of your report shows the pedigree of the horse being analyzed with the individual algorithm score of each variable including the inbreeding coefficient.</p>")
    pdf.write_html("<p line-height='1.3'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;The next section are the top sires sorted by the <i>Equi-Source Score</i> and compared to multiple constants in your horse's pedigree using a tiering system. Tier 1 being the most favorable constant and tier 4 being the least favorable depending on the location in the pedigree.</p>")

    ################# page 4 #################
    pdf.add_page()
    pdf.ln()
    pdf.ln()

    pdf.set_font('Times', 'B', 15)
    pdf.set_text_color(0, 50, 120)
    pdf.cell(lmargin)
    pdf.cell(w=0, h=15, text="Interpreting the data", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.ln()

    pdf.set_font('Helvetica', '', 12)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(lmargin)
    pdf.write_html("<p line-height='1.3'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;While statistics can provide insight in helping you make decisions, there are many other variables to consider such as your overall goals. Do you want a prospect you can win on right now, or do you want to build the foundation to a successful breeding program? These are the questions that Equi-Source can help you answer and determine where each individual horse works best in your program. Just because a particular cross results in a lower grade, does not mean it won't be successful. It's imperative that industry professionals continue to breed the horses they believe in regardless of what the data illustrates, because it is impossible to capture every successful variable no matter how powerful the technology. The more data we compile, the more accurate the results and future iterations will reflect that movement.</p>")

    ################# page (Pedigree Table) #################
    pdf.add_page()
    pdf.set_line_width(2)
    pdf.set_fill_color(r=255, g=255, b=255)
    pdf.rect(x=50, y=80, w=280, h=70, style="D")
    pdf.rect(x=450, y=80, w=100, h=70, style="D")

    pdf.ln()
    pdf.ln()
    pdf.set_font('Times', '', 25)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(lmargin+10)
    pdf.cell(w=280, h=30, text=sheetName, new_x=XPos.LMARGIN, new_y=YPos.TOP)

    pdf.set_font('Times', '', 60)
    pdf.set_text_color(grade_color[0], grade_color[1], grade_color[2])
    pdf.cell(420)
    pdf.cell(w=100, h=40, text=letter_grade, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_font('Times', '', 18)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(lmargin+10)
    pdf.cell(w=280, h=10, text=f"{pedigree_dict['birth']} {pedigree_dict['sex']}", new_x=XPos.LMARGIN, new_y=YPos.TOP)

    pdf.set_font('Times', '', 10)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(420)
    pdf.cell(w=100, h=25, text=f"VARIANT = {v_sum}", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_line_width(0.5)
    pdf.ln(150)
    pdf.set_fill_color(255, 0, 0) ## It's just for red block square
    # MMM
    anc_mmm = pedigree_dict["pedigree"][0]
    pdf.cell(280)
    if "*" in anc_mmm:
        pdf.cell(w=120, h=0, text=anc_mmm.replace("*",""), align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=310, y=290, w=120, h=25, style="D")
        pdf.set_draw_color(255, 0, 0)
        pdf.rect(x=425, y=290, w=5, h=25, style="DF")
    else:
        pdf.cell(w=120, h=0, text=anc_mmm, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=310, y=290, w=120, h=25, style="D")
    # MMMM
    anc_mmmm = pedigree_dict["pedigree"][1]
    pdf.cell(410)
    if "*" in anc_mmmm:
        pdf.cell(w=120, h=0, text=anc_mmmm.replace("*",""), align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=440, y=290, w=120, h=25, style="D")
        pdf.set_draw_color(255, 0, 0)
        pdf.rect(x=555, y=290, w=5, h=25, style="DF")
    else:
        pdf.cell(w=120, h=0, text=anc_mmmm, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=440, y=290, w=120, h=25, style="D")
    pdf.ln(20)
    # MM
    anc_mm = pedigree_dict["pedigree"][2]
    pdf.cell(150)
    if "*" in anc_mm:
        pdf.cell(w=120, h=0, text=anc_mm.replace("*",""), align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=180, y=310, w=120, h=25, style="D")
        pdf.set_draw_color(255, 0, 0)
        pdf.rect(x=295, y=310, w=5, h=25, style="DF")
    else:
        pdf.cell(w=120, h=0, text=anc_mm, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=180, y=310, w=120, h=25, style="D")
    pdf.ln(20)
    # MMF
    anc_mmf = pedigree_dict["pedigree"][3]
    pdf.cell(280)
    if "*" in anc_mmf:
        pdf.cell(w=120, h=0, text=anc_mmf.replace("*",""), align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=310, y=330, w=120, h=25, style="D")
        pdf.set_draw_color(255, 0, 0)
        pdf.rect(x=425, y=330, w=5, h=25, style="DF")
    else:
        pdf.cell(w=120, h=0, text=anc_mmf, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=310, y=330, w=120, h=25, style="D")
    # MMFM
    anc_mmfm = pedigree_dict["pedigree"][4]
    pdf.cell(410)
    if "*" in anc_mmfm:
        pdf.cell(w=120, h=0, text=anc_mmfm.replace("*",""), align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=440, y=330, w=120, h=25, style="D")
        pdf.set_draw_color(255, 0, 0)
        pdf.rect(x=555, y=330, w=5, h=25, style="DF")
    else:
        pdf.cell(w=120, h=0, text=anc_mmfm, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=440, y=330, w=120, h=25, style="D")
    pdf.ln(20)
    # M
    anc_m = pedigree_dict["pedigree"][5]
    pdf.cell(20)
    if "*" in anc_m:
        pdf.cell(w=120, h=0, text=anc_m.replace("*",""), align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=50, y=350, w=120, h=25, style="D")
        pdf.set_draw_color(255, 0, 0)
        pdf.rect(x=165, y=350, w=5, h=25, style="DF")
    else:
        pdf.cell(w=120, h=0, text=anc_m, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=50, y=350, w=120, h=25, style="D")
    pdf.ln(20)
    # MFM
    anc_mfm = pedigree_dict["pedigree"][6]
    pdf.cell(280)
    if "*" in anc_mfm:
        pdf.cell(w=120, h=0, text=anc_mfm.replace("*",""), align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=310, y=370, w=120, h=25, style="D")
        pdf.set_draw_color(255, 0, 0)
        pdf.rect(x=425, y=370, w=5, h=25, style="DF")
    else:
        pdf.cell(w=120, h=0, text=anc_mfm, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=310, y=370, w=120, h=25, style="D")
    # MFMM
    anc_mfmm = pedigree_dict["pedigree"][7]
    pdf.cell(410)
    if "*" in anc_mfmm:
        pdf.cell(w=120, h=0, text=anc_mfmm.replace("*",""), align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=440, y=370, w=120, h=25, style="D")
        pdf.set_draw_color(255, 0, 0)
        pdf.rect(x=555, y=370, w=5, h=25, style="DF")
    else:
        pdf.cell(w=120, h=0, text=anc_mfmm, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=440, y=370, w=120, h=25, style="D")
    pdf.ln(20)
    # MF
    anc_mf = pedigree_dict["pedigree"][8]
    pdf.cell(150)
    if "*" in anc_mf:
        pdf.cell(w=120, h=0, text=anc_mf.replace("*",""), align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=180, y=390, w=120, h=25, style="D")
        pdf.set_draw_color(255, 0, 0)
        pdf.rect(x=295, y=390, w=5, h=25, style="DF")
    else:
        pdf.cell(w=120, h=0, text=anc_mf, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=180, y=390, w=120, h=25, style="D")
    pdf.ln(20)
    # MFF
    anc_mff = pedigree_dict["pedigree"][9]
    pdf.cell(280)
    if "*" in anc_mff:
        pdf.cell(w=120, h=0, text=anc_mff.replace("*",""), align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=310, y=410, w=120, h=25, style="D")
        pdf.set_draw_color(255, 0, 0)
        pdf.rect(x=425, y=410, w=5, h=25, style="DF")
    else:
        pdf.cell(w=120, h=0, text=anc_mff, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=310, y=410, w=120, h=25, style="D")
    # MFFM
    anc_mffm = pedigree_dict["pedigree"][10]
    pdf.cell(410)
    if "*" in anc_mffm:
        pdf.cell(w=120, h=0, text=anc_mffm.replace("*",""), align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=440, y=410, w=120, h=25, style="D")
        pdf.set_draw_color(255, 0, 0)
        pdf.rect(x=555, y=410, w=5, h=25, style="DF")
    else:
        pdf.cell(w=120, h=0, text=anc_mffm, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=440, y=410, w=120, h=25, style="D")
    pdf.ln(40)
    # FMM
    anc_fmm = pedigree_dict["pedigree"][11]
    pdf.cell(280)
    if "*" in anc_fmm:
        pdf.cell(w=120, h=0, text=anc_fmm.replace("*",""), align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=310, y=450, w=120, h=25, style="D")
        pdf.set_draw_color(255, 0, 0)
        pdf.rect(x=425, y=450, w=5, h=25, style="DF")
    else:
        pdf.cell(w=120, h=0, text=anc_fmm, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=310, y=450, w=120, h=25, style="D")
    # FMMM
    anc_fmmm = pedigree_dict["pedigree"][12]
    pdf.cell(410)
    if "*" in anc_fmmm:
        pdf.cell(w=120, h=0, text=anc_fmmm.replace("*",""), align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=440, y=450, w=120, h=25, style="D")
        pdf.set_draw_color(255, 0, 0)
        pdf.rect(x=555, y=450, w=5, h=25, style="DF")
    else:
        pdf.cell(w=120, h=0, text=anc_fmmm, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=440, y=450, w=120, h=25, style="D")
    pdf.ln(20)
    # FM
    anc_fm = pedigree_dict["pedigree"][13]
    pdf.cell(150)
    if "*" in anc_fm:
        pdf.cell(w=120, h=0, text=anc_fm.replace("*",""), align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=180, y=470, w=120, h=25, style="D")
        pdf.set_draw_color(255, 0, 0)
        pdf.rect(x=295, y=470, w=5, h=25, style="DF")
    else:
        pdf.cell(w=120, h=0, text=anc_fm, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=180, y=470, w=120, h=25, style="D")
    pdf.ln(20)
    # FMF
    anc_fmf = pedigree_dict["pedigree"][14]
    pdf.cell(280)
    if "*" in anc_fmf:
        pdf.cell(w=120, h=0, text=anc_fmf.replace("*",""), align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=310, y=490, w=120, h=25, style="D")
        pdf.set_draw_color(255, 0, 0)
        pdf.rect(x=425, y=490, w=5, h=25, style="DF")
    else:
        pdf.cell(w=120, h=0, text=anc_fmf, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=310, y=490, w=120, h=25, style="D")
    # FMFM
    anc_fmfm = pedigree_dict["pedigree"][15]
    pdf.cell(410)
    if "*" in anc_fmfm:
        pdf.cell(w=120, h=0, text=anc_fmfm.replace("*",""), align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=440, y=490, w=120, h=25, style="D")
        pdf.set_draw_color(255, 0, 0)
        pdf.rect(x=555, y=490, w=5, h=25, style="DF")
    else:
        pdf.cell(w=120, h=0, text=anc_fmfm, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=440, y=490, w=120, h=25, style="D")
    pdf.ln(20)
    # F
    anc_f = pedigree_dict["pedigree"][16]
    pdf.cell(20)
    if "*" in anc_f:
        pdf.cell(w=120, h=0, text=anc_f.replace("*",""), align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=50, y=510, w=120, h=25, style="D")
        pdf.set_draw_color(255, 0, 0)
        pdf.rect(x=165, y=510, w=5, h=25, style="DF")
    else:
        pdf.cell(w=120, h=0, text=anc_f, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=50, y=510, w=120, h=25, style="D")
    pdf.ln(20)
    # FFM
    anc_ffm = pedigree_dict["pedigree"][17]
    pdf.cell(280)
    if "*" in anc_ffm:
        pdf.cell(w=120, h=0, text=anc_ffm.replace("*",""), align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=310, y=530, w=120, h=25, style="D")
        pdf.set_draw_color(255, 0, 0)
        pdf.rect(x=425, y=530, w=5, h=25, style="DF")
    else:
        pdf.cell(w=120, h=0, text=anc_ffm, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=310, y=530, w=120, h=25, style="D")
    # FFMM
    anc_ffmm = pedigree_dict["pedigree"][18]
    pdf.cell(410)
    if "*" in anc_ffmm:
        pdf.cell(w=120, h=0, text=anc_ffmm.replace("*",""), align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=440, y=530, w=120, h=25, style="D")
        pdf.set_draw_color(255, 0, 0)
        pdf.rect(x=555, y=530, w=5, h=25, style="DF")
    else:
        pdf.cell(w=120, h=0, text=anc_ffmm, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=440, y=530, w=120, h=25, style="D")
    pdf.ln(20)
    # FF
    anc_ff = pedigree_dict["pedigree"][19]
    pdf.cell(150)
    if "*" in anc_ff:
        pdf.cell(w=120, h=0, text=anc_ff.replace("*",""), align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=180, y=550, w=120, h=25, style="D")
        pdf.set_draw_color(255, 0, 0)
        pdf.rect(x=295, y=550, w=5, h=25, style="DF")
    else:
        pdf.cell(w=120, h=0, text=anc_ff, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=180, y=550, w=120, h=25, style="D")
    pdf.ln(20)
    # FFF
    anc_fff = pedigree_dict["pedigree"][20]
    pdf.cell(280)
    if "*" in anc_fff:
        pdf.cell(w=120, h=0, text=anc_fff.replace("*",""), align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=310, y=570, w=120, h=25, style="D")
        pdf.set_draw_color(255, 0, 0)
        pdf.rect(x=425, y=570, w=5, h=25, style="DF")
    else:
        pdf.cell(w=120, h=0, text=anc_fff, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=310, y=570, w=120, h=25, style="D")
    # FFFM
    anc_fffm = pedigree_dict["pedigree"][21]
    pdf.cell(410)
    if "*" in anc_fffm:
        pdf.cell(w=120, h=0, text=anc_fffm.replace("*",""), align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=440, y=570, w=120, h=25, style="D")
        pdf.set_draw_color(255, 0, 0)
        pdf.rect(x=555, y=570, w=5, h=25, style="DF")
    else:
        pdf.cell(w=120, h=0, text=anc_fffm, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(x=440, y=570, w=120, h=25, style="D")

    pdf.set_fill_color(255, 255, 255) # Back to white background
    pdf.ln(70)

    pdf.set_left_margin(70)
    TABLE_DATA = (
        ("Sire", "Dam's Sire", "2nd Dam's Sire", "3rd Dam's Sire", "Inbreeding Coefficient"),
        (get2DigitsStringValue(v_sire), get2DigitsStringValue(v_damssire), get2DigitsStringValue(v_damssire2), get2DigitsStringValue(v_damssire3), coi_val),
    )

    with pdf.table(text_align=Align.C, col_widths=100, line_height=10, padding=2) as table:
        for data_row in TABLE_DATA:
            row = table.row()
            for datum in data_row:
                row.cell(datum, padding=(8, 5, 8, 5))
    pdf.set_left_margin(30)

    ################# page (Tier 1 Suggestions sorted by Variant) #################
    if len(tier1_sugs) != 0:
        sorted_tier1_sugs = sortByVariant(tier1_sugs, genType)
        for i in range(math.ceil(len(sorted_tier1_sugs) / 10)):
            page_label = "TOP STALLIONS BY EQUI-SOURCE SCORE"
            if i != 0:
                page_label += "(CONTINUED)"
            pdf.add_page()
            pdf.set_line_width(2)
            pdf.set_fill_color(r=255, g=255, b=255)
            pdf.rect(x=50, y=80, w=280, h=70, style="D")
            pdf.rect(x=450, y=80, w=100, h=70, style="D")

            pdf.ln()
            pdf.ln()
            pdf.set_font('Times', '', 25)
            pdf.set_text_color(0, 0, 0)
            pdf.cell(lmargin+10)
            pdf.cell(w=0, h=30, text="Tier 1", new_x=XPos.LMARGIN, new_y=YPos.TOP)

            pdf.set_font('Times', '', 60)
            pdf.cell(420)
            pdf.cell(w=100, h=40, text=f"{len(tier1_sugs)}", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.set_font('Times', '', 18)
            pdf.cell(lmargin+10)
            pdf.cell(w=0, h=10, text=f"{tier1_filter_label}", new_x=XPos.LMARGIN, new_y=YPos.TOP)

            pdf.set_font('Times', '', 10)
            pdf.cell(420)
            pdf.cell(w=100, h=25, text="MATCHES FOUND", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.ln(100)

            pdf.set_font('Times', '', 18)
            pdf.cell(lmargin)
            pdf.cell(w=0, h=0, text=page_label, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.ln(100)
            pdf.set_left_margin(70)
            pdf.set_line_width(0.5)
            pdf.set_font('Times', '', 10)

            TABLE_HEADER_DATA = [
                ["Stallion", "1D Rate", "Variant", "Equi-Source Score", "Inbreeding Coefficient of foal"]
            ]
            TABLE_DATA = TABLE_HEADER_DATA + sorted_tier1_sugs[i*10:i*10+10]

            with pdf.table(text_align=Align.C, col_widths=100, line_height=10, padding=2) as table:
                for data_row in TABLE_DATA:
                    row = table.row()
                    for datum in data_row:
                        row.cell(datum, padding=(8, 5, 8, 5))
            pdf.set_left_margin(30)

    ################# page (Tier 1 Suggestions sorted by 1D Rate) #################
        sorted_tier1_sugs = sortByRate(tier1_sugs, genType)
        for i in range(math.ceil(len(sorted_tier1_sugs) / 10)):
            page_label = "TOP STALLIONS BY 1D RATE"
            if i != 0:
                page_label += "(CONTINUED)"
            pdf.add_page()
            pdf.set_line_width(2)
            pdf.set_fill_color(r=255, g=255, b=255)
            pdf.rect(x=50, y=80, w=280, h=70, style="D")
            pdf.rect(x=450, y=80, w=100, h=70, style="D")

            pdf.ln()
            pdf.ln()
            pdf.set_font('Times', '', 25)
            pdf.set_text_color(0, 0, 0)
            pdf.cell(lmargin+10)
            pdf.cell(w=0, h=30, text="Tier 1", new_x=XPos.LMARGIN, new_y=YPos.TOP)

            pdf.set_font('Times', '', 60)
            pdf.cell(420)
            pdf.cell(w=100, h=40, text=f"{len(tier1_sugs)}", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.set_font('Times', '', 18)
            pdf.cell(lmargin+10)
            pdf.cell(w=0, h=10, text=f"{tier1_filter_label}", new_x=XPos.LMARGIN, new_y=YPos.TOP)

            pdf.set_font('Times', '', 10)
            pdf.cell(420)
            pdf.cell(w=100, h=25, text="MATCHES FOUND", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.ln(100)

            pdf.set_font('Times', '', 18)
            pdf.cell(lmargin)
            pdf.cell(w=0, h=0, text=page_label, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.ln(100)
            pdf.set_left_margin(70)
            pdf.set_line_width(0.5)
            pdf.set_font('Times', '', 10)
            TABLE_HEADER_DATA = [
                ["Stallion", "1D Rate", "Variant", "Equi-Source Score", "Inbreeding Coefficient of foal"]
            ]
            TABLE_DATA = TABLE_HEADER_DATA + sorted_tier1_sugs[i*10:i*10+10]

            with pdf.table(text_align=Align.C, col_widths=100, line_height=10, padding=2) as table:
                for data_row in TABLE_DATA:
                    row = table.row()
                    for datum in data_row:
                        row.cell(datum, padding=(8, 5, 8, 5))

            pdf.set_left_margin(30)

    ################# page (Tier 1 Suggestions sorted by Inbreeding Coefficient) #################
        sorted_tier1_sugs = sortByCoi(tier1_sugs, genType)
        for i in range(math.ceil(len(sorted_tier1_sugs) / 10)):
            page_label = "TOP STALLIONS BY INBREEDING COEFFICIENT"
            if i != 0:
                page_label += "(CONTINUED)"
            pdf.add_page()
            pdf.set_line_width(2)
            pdf.set_fill_color(r=255, g=255, b=255)
            pdf.rect(x=50, y=80, w=280, h=70, style="D")
            pdf.rect(x=450, y=80, w=100, h=70, style="D")

            pdf.ln()
            pdf.ln()
            pdf.set_font('Times', '', 25)
            pdf.set_text_color(0, 0, 0)
            pdf.cell(lmargin+10)
            pdf.cell(w=0, h=30, text="Tier 1", new_x=XPos.LMARGIN, new_y=YPos.TOP)

            pdf.set_font('Times', '', 60)
            pdf.cell(420)
            pdf.cell(w=100, h=40, text=f"{len(tier1_sugs)}", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.set_font('Times', '', 18)
            pdf.cell(lmargin+10)
            pdf.cell(w=0, h=10, text=f"{tier1_filter_label}", new_x=XPos.LMARGIN, new_y=YPos.TOP)

            pdf.set_font('Times', '', 10)
            pdf.cell(420)
            pdf.cell(w=100, h=25, text="MATCHES FOUND", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.ln(100)

            pdf.set_font('Times', '', 18)
            pdf.cell(lmargin)
            pdf.cell(w=0, h=0, text=page_label, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.ln(100)
            pdf.set_left_margin(70)
            pdf.set_line_width(0.5)
            pdf.set_font('Times', '', 10)
            TABLE_HEADER_DATA = [
                ["Stallion", "1D Rate", "Variant", "Equi-Source Score", "Inbreeding Coefficient of foal"]
            ]
            TABLE_DATA = TABLE_HEADER_DATA + sorted_tier1_sugs[i*10:i*10+10]

            with pdf.table(text_align=Align.C, col_widths=100, line_height=10, padding=2) as table:
                for data_row in TABLE_DATA:
                    row = table.row()
                    for datum in data_row:
                        row.cell(datum, padding=(8, 5, 8, 5))

            pdf.set_left_margin(30)

    ################# page (Tier 2 Suggestions sorted by Variant) #################
    if len(tier2_sugs):
        sorted_tier2_sugs = sortByVariant(tier2_sugs, genType)
        for i in range(math.ceil(len(sorted_tier2_sugs) / 10)):
            page_label = "TOP STALLIONS BY EQUI-SOURCE SCORE"
            if i != 0:
                page_label += "(CONTINUED)"
            pdf.add_page()
            pdf.set_line_width(2)
            pdf.set_fill_color(r=255, g=255, b=255)
            pdf.rect(x=50, y=80, w=280, h=70, style="D")
            pdf.rect(x=450, y=80, w=100, h=70, style="D")

            pdf.ln()
            pdf.ln()
            pdf.set_font('Times', '', 25)
            pdf.set_text_color(0, 0, 0)
            pdf.cell(lmargin+10)
            pdf.cell(w=0, h=30, text="Tier 2", new_x=XPos.LMARGIN, new_y=YPos.TOP)

            pdf.set_font('Times', '', 60)
            pdf.cell(420)
            pdf.cell(w=100, h=40, text=f"{len(tier2_sugs)}", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.set_font('Times', '', 18)
            pdf.cell(lmargin+10)
            pdf.cell(w=0, h=10, text="Stallion Alternative", new_x=XPos.LMARGIN, new_y=YPos.TOP)

            pdf.set_font('Times', '', 10)
            pdf.cell(420)
            pdf.cell(w=100, h=25, text="MATCHES FOUND", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.ln(100)

            pdf.set_font('Times', '', 18)
            pdf.cell(lmargin)
            pdf.cell(w=0, h=0, text=page_label, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.ln(100)
            pdf.set_left_margin(70)
            pdf.set_line_width(0.5)
            pdf.set_font('Times', '', 10)
            TABLE_HEADER_DATA = [
                ["Stallion", "1D Rate", "Variant", "Equi-Source Score", "Inbreeding Coefficient of foal"]
            ]
            TABLE_DATA = TABLE_HEADER_DATA + sorted_tier2_sugs[i*10:i*10+10]

            with pdf.table(text_align=Align.C, col_widths=100, line_height=10, padding=2) as table:
                for data_row in TABLE_DATA:
                    row = table.row()
                    for datum in data_row:
                        row.cell(datum, padding=(8, 5, 8, 5))

            pdf.set_left_margin(30)

    ################# page (Tier 2 Suggestions sorted by 1D Rate) #################
        sorted_tier2_sugs = sortByRate(tier2_sugs, genType)
        for i in range(math.ceil(len(sorted_tier2_sugs) / 10)):
            page_label = "TOP STALLIONS BY 1D RATE"
            if i != 0:
                page_label += "(CONTINUED)"
            pdf.add_page()
            pdf.set_line_width(2)
            pdf.set_fill_color(r=255, g=255, b=255)
            pdf.rect(x=50, y=80, w=280, h=70, style="D")
            pdf.rect(x=450, y=80, w=100, h=70, style="D")

            pdf.ln()
            pdf.ln()
            pdf.set_font('Times', '', 25)
            pdf.set_text_color(0, 0, 0)
            pdf.cell(lmargin+10)
            pdf.cell(w=0, h=30, text="Tier 2", new_x=XPos.LMARGIN, new_y=YPos.TOP)

            pdf.set_font('Times', '', 60)
            pdf.cell(420)
            pdf.cell(w=100, h=40, text=f"{len(tier2_sugs)}", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.set_font('Times', '', 18)
            pdf.cell(lmargin+10)
            pdf.cell(w=0, h=10, text="Stallion Alternative", new_x=XPos.LMARGIN, new_y=YPos.TOP)

            pdf.set_font('Times', '', 10)
            pdf.cell(420)
            pdf.cell(w=100, h=25, text="MATCHES FOUND", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.ln(100)

            pdf.set_font('Times', '', 18)
            pdf.cell(lmargin)
            pdf.cell(w=0, h=0, text=page_label, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.ln(100)
            pdf.set_left_margin(70)
            pdf.set_line_width(0.5)
            pdf.set_font('Times', '', 10)
            TABLE_HEADER_DATA = [
                ["Stallion", "1D Rate", "Variant", "Equi-Source Score", "Inbreeding Coefficient of foal"]
            ]
            TABLE_DATA = TABLE_HEADER_DATA + sorted_tier2_sugs[i*10:i*10+10]

            with pdf.table(text_align=Align.C, col_widths=100, line_height=10, padding=2) as table:
                for data_row in TABLE_DATA:
                    row = table.row()
                    for datum in data_row:
                        row.cell(datum, padding=(8, 5, 8, 5))

            pdf.set_left_margin(30)

    ################# page (Tier 2 Suggestions sorted by Inbreeding Coefficient) #################
        sorted_tier2_sugs = sortByCoi(tier2_sugs, genType)
        for i in range(math.ceil(len(sorted_tier2_sugs) / 10)):
            page_label = "TOP STALLIONS BY INBREEDING COEFFICIENT"
            if i != 0:
                page_label += "(CONTINUED)"
            pdf.add_page()
            pdf.set_line_width(2)
            pdf.set_fill_color(r=255, g=255, b=255)
            pdf.rect(x=50, y=80, w=280, h=70, style="D")
            pdf.rect(x=450, y=80, w=100, h=70, style="D")

            pdf.ln()
            pdf.ln()
            pdf.set_font('Times', '', 25)
            pdf.set_text_color(0, 0, 0)
            pdf.cell(lmargin+10)
            pdf.cell(w=0, h=30, text="Tier 2", new_x=XPos.LMARGIN, new_y=YPos.TOP)

            pdf.set_font('Times', '', 60)
            pdf.cell(420)
            pdf.cell(w=100, h=40, text=f"{len(tier2_sugs)}", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.set_font('Times', '', 18)
            pdf.cell(lmargin+10)
            pdf.cell(w=0, h=10, text="Stallion Alternative", new_x=XPos.LMARGIN, new_y=YPos.TOP)

            pdf.set_font('Times', '', 10)
            pdf.cell(420)
            pdf.cell(w=100, h=25, text="MATCHES FOUND", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.ln(100)

            pdf.set_font('Times', '', 18)
            pdf.cell(lmargin)
            pdf.cell(w=0, h=0, text=page_label, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.ln(100)
            pdf.set_left_margin(70)
            pdf.set_line_width(0.5)
            pdf.set_font('Times', '', 10)
            TABLE_HEADER_DATA = [
                ["Stallion", "1D Rate", "Variant", "Equi-Source Score", "Inbreeding Coefficient of foal"]
            ]
            TABLE_DATA = TABLE_HEADER_DATA + sorted_tier2_sugs[i*10:i*10+10]

            with pdf.table(text_align=Align.C, col_widths=100, line_height=10, padding=2) as table:
                for data_row in TABLE_DATA:
                    row = table.row()
                    for datum in data_row:
                        row.cell(datum, padding=(8, 5, 8, 5))

            pdf.set_left_margin(30)

    ################# page (Tier 3 Suggestions sorted by Variant) #################
    if len(tier3_sugs) != 0:
        sorted_tier3_sugs = sortByVariant(tier3_sugs, genType)
        for i in range(math.ceil(len(sorted_tier3_sugs) / 10)):
            page_label = "TOP STALLIONS BY EQUI-SOURCE SCORE"
            if i != 0:
                page_label += "(CONTINUED)"
            pdf.add_page()
            pdf.set_line_width(2)
            pdf.set_fill_color(r=255, g=255, b=255)
            pdf.rect(x=50, y=80, w=280, h=70, style="D")
            pdf.rect(x=450, y=80, w=100, h=70, style="D")

            pdf.ln()
            pdf.ln()
            pdf.set_font('Times', '', 25)
            pdf.set_text_color(0, 0, 0)
            pdf.cell(lmargin+10)
            pdf.cell(w=0, h=30, text="Tier 3", new_x=XPos.LMARGIN, new_y=YPos.TOP)

            pdf.set_font('Times', '', 60)
            pdf.cell(420)
            pdf.cell(w=100, h=40, text=f"{len(tier3_sugs)}", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.set_font('Times', '', 18)
            pdf.cell(lmargin+10)
            pdf.cell(w=0, h=10, text=f"{tier3_filter_label}", new_x=XPos.LMARGIN, new_y=YPos.TOP)

            pdf.set_font('Times', '', 10)
            pdf.cell(420)
            pdf.cell(w=100, h=25, text="MATCHES FOUND", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.ln(100)

            pdf.set_font('Times', '', 18)
            pdf.cell(lmargin)
            pdf.cell(w=0, h=0, text=page_label, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.ln(100)
            pdf.set_left_margin(70)
            pdf.set_line_width(0.5)
            pdf.set_font('Times', '', 10)
            TABLE_HEADER_DATA = [
                ["Stallion", "1D Rate", "Variant", "Equi-Source Score", "Inbreeding Coefficient of foal"]
            ]
            TABLE_DATA = TABLE_HEADER_DATA + sorted_tier3_sugs[i*10:i*10+10]

            with pdf.table(text_align=Align.C, col_widths=100, line_height=10, padding=2) as table:
                for data_row in TABLE_DATA:
                    row = table.row()
                    for datum in data_row:
                        row.cell(datum, padding=(8, 5, 8, 5))

            pdf.set_left_margin(30)

    ################# page (Tier 3 Suggestions sorted by 1D Rate) #################
        sorted_tier3_sugs = sortByRate(tier3_sugs, genType)
        for i in range(math.ceil(len(sorted_tier3_sugs) / 10)):
            page_label = "TOP STALLIONS BY 1D RATE"
            if i != 0:
                page_label += "(CONTINUED)"
            pdf.add_page()
            pdf.set_line_width(2)
            pdf.set_fill_color(r=255, g=255, b=255)
            pdf.rect(x=50, y=80, w=280, h=70, style="D")
            pdf.rect(x=450, y=80, w=100, h=70, style="D")

            pdf.ln()
            pdf.ln()
            pdf.set_font('Times', '', 25)
            pdf.set_text_color(0, 0, 0)
            pdf.cell(lmargin+10)
            pdf.cell(w=0, h=30, text="Tier 3", new_x=XPos.LMARGIN, new_y=YPos.TOP)

            pdf.set_font('Times', '', 60)
            pdf.cell(420)
            pdf.cell(w=100, h=40, text=f"{len(tier3_sugs)}", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.set_font('Times', '', 18)
            pdf.cell(lmargin+10)
            pdf.cell(w=0, h=10, text=f"{tier3_filter_label}", new_x=XPos.LMARGIN, new_y=YPos.TOP)

            pdf.set_font('Times', '', 10)
            pdf.cell(420)
            pdf.cell(w=100, h=25, text="MATCHES FOUND", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.ln(100)

            pdf.set_font('Times', '', 18)
            pdf.cell(lmargin)
            pdf.cell(w=0, h=0, text=page_label, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.ln(100)
            pdf.set_left_margin(70)
            pdf.set_line_width(0.5)
            pdf.set_font('Times', '', 10)
            TABLE_HEADER_DATA = [
                ["Stallion", "1D Rate", "Variant", "Equi-Source Score", "Inbreeding Coefficient of foal"]
            ]
            TABLE_DATA = TABLE_HEADER_DATA + sorted_tier3_sugs[i*10:i*10+10]

            with pdf.table(text_align=Align.C, col_widths=100, line_height=10, padding=2) as table:
                for data_row in TABLE_DATA:
                    row = table.row()
                    for datum in data_row:
                        row.cell(datum, padding=(8, 5, 8, 5))

            pdf.set_left_margin(30)

    ################# page (Tier 3 Suggestions sorted by Inbreeding Coefficient) #################
        sorted_tier3_sugs = sortByCoi(tier3_sugs, genType)
        for i in range(math.ceil(len(sorted_tier3_sugs) / 10)):
            page_label = "TOP STALLIONS BY INBREEDING COEFFICIENT"
            if i != 0:
                page_label += "(CONTINUED)"
            pdf.add_page()
            pdf.set_line_width(2)
            pdf.set_fill_color(r=255, g=255, b=255)
            pdf.rect(x=50, y=80, w=280, h=70, style="D")
            pdf.rect(x=450, y=80, w=100, h=70, style="D")

            pdf.ln()
            pdf.ln()
            pdf.set_font('Times', '', 25)
            pdf.set_text_color(0, 0, 0)
            pdf.cell(lmargin+10)
            pdf.cell(w=0, h=30, text="Tier 3", new_x=XPos.LMARGIN, new_y=YPos.TOP)

            pdf.set_font('Times', '', 60)
            pdf.cell(420)
            pdf.cell(w=100, h=40, text=f"{len(tier3_sugs)}", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.set_font('Times', '', 18)
            pdf.cell(lmargin+10)
            pdf.cell(w=0, h=10, text=f"{tier3_filter_label}", new_x=XPos.LMARGIN, new_y=YPos.TOP)

            pdf.set_font('Times', '', 10)
            pdf.cell(420)
            pdf.cell(w=100, h=25, text="MATCHES FOUND", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.ln(100)

            pdf.set_font('Times', '', 18)
            pdf.cell(lmargin)
            pdf.cell(w=0, h=0, text=page_label, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.ln(100)
            pdf.set_left_margin(70)
            pdf.set_line_width(0.5)
            pdf.set_font('Times', '', 10)
            TABLE_HEADER_DATA = [
                ["Stallion", "1D Rate", "Variant", "Equi-Source Score", "Inbreeding Coefficient of foal"]
            ]
            TABLE_DATA = TABLE_HEADER_DATA + sorted_tier3_sugs[i*10:i*10+10]

            with pdf.table(text_align=Align.C, col_widths=100, line_height=10, padding=2) as table:
                for data_row in TABLE_DATA:
                    row = table.row()
                    for datum in data_row:
                        row.cell(datum, padding=(8, 5, 8, 5))

            pdf.set_left_margin(30)

    ################# page (Tier 4 Suggestions sorted by Variant) #################
    if len(tier4_sugs) != 0:
        sorted_tier4_sugs = sortByVariant(tier4_sugs, genType)
        for i in range(math.ceil(len(sorted_tier4_sugs) / 10)):
            page_label = "TOP STALLIONS BY EQUI-SOURCE SCORE"
            if i != 0:
                page_label += "(CONTINUED)"
            pdf.add_page()
            pdf.set_line_width(2)
            pdf.set_fill_color(r=255, g=255, b=255)
            pdf.rect(x=50, y=80, w=280, h=90, style="D")
            pdf.rect(x=450, y=80, w=100, h=70, style="D")

            pdf.ln()
            pdf.ln()
            pdf.set_font('Times', '', 25)
            pdf.set_text_color(0, 0, 0)
            pdf.cell(lmargin+10)
            pdf.cell(w=0, h=30, text="Tier 4", new_x=XPos.LMARGIN, new_y=YPos.TOP)

            pdf.set_font('Times', '', 60)
            pdf.cell(420)
            pdf.cell(w=100, h=40, text=f"{len(tier4_sugs)}", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.set_font('Times', '', 18)
            pdf.cell(lmargin+10)
            pdf.multi_cell(w=260, h=20, text=f"{tier4_filter_label}", new_x=XPos.LMARGIN, new_y=YPos.TOP)

            pdf.set_font('Times', '', 10)
            pdf.cell(420)
            pdf.cell(w=100, h=25, text="MATCHES FOUND", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.ln(100)

            pdf.set_font('Times', '', 18)
            pdf.cell(lmargin)
            pdf.cell(w=0, h=0, text=page_label, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.ln(100)
            pdf.set_left_margin(70)
            pdf.set_line_width(0.5)
            pdf.set_font('Times', '', 10)
            TABLE_HEADER_DATA = [
                ["Stallion", "1D Rate", "Variant", "Equi-Source Score", "Inbreeding Coefficient of foal"]
            ]
            TABLE_DATA = TABLE_HEADER_DATA + sorted_tier4_sugs[i*10:i*10+10]

            with pdf.table(text_align=Align.C, col_widths=100, line_height=10, padding=2) as table:
                for data_row in TABLE_DATA:
                    row = table.row()
                    for datum in data_row:
                        row.cell(datum, padding=(8, 5, 8, 5))

            pdf.set_left_margin(30)

    ################# page (Tier 4 Suggestions sorted by 1D Rate) #################
        sorted_tier4_sugs = sortByRate(tier4_sugs, genType)
        for i in range(math.ceil(len(sorted_tier4_sugs) / 10)):
            page_label = "TOP STALLIONS BY 1D RATE"
            if i != 0:
                page_label += "(CONTINUED)"
            pdf.add_page()
            pdf.set_line_width(2)
            pdf.set_fill_color(r=255, g=255, b=255)
            pdf.rect(x=50, y=80, w=280, h=90, style="D")
            pdf.rect(x=450, y=80, w=100, h=70, style="D")

            pdf.ln()
            pdf.ln()
            pdf.set_font('Times', '', 25)
            pdf.set_text_color(0, 0, 0)
            pdf.cell(lmargin+10)
            pdf.cell(w=0, h=30, text="Tier 4", new_x=XPos.LMARGIN, new_y=YPos.TOP)

            pdf.set_font('Times', '', 60)
            pdf.cell(420)
            pdf.cell(w=100, h=40, text=f"{len(tier4_sugs)}", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.set_font('Times', '', 18)
            pdf.cell(lmargin+10)
            pdf.multi_cell(w=260, h=20, text=f"{tier4_filter_label}", new_x=XPos.LMARGIN, new_y=YPos.TOP)

            pdf.set_font('Times', '', 10)
            pdf.cell(420)
            pdf.cell(w=100, h=25, text="MATCHES FOUND", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.ln(100)

            pdf.set_font('Times', '', 18)
            pdf.cell(lmargin)
            pdf.cell(w=0, h=0, text=page_label, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.ln(100)
            pdf.set_left_margin(70)
            pdf.set_line_width(0.5)
            pdf.set_font('Times', '', 10)
            TABLE_HEADER_DATA = [
                ["Stallion", "1D Rate", "Variant", "Equi-Source Score", "Inbreeding Coefficient of foal"]
            ]
            TABLE_DATA = TABLE_HEADER_DATA + sorted_tier4_sugs[i*10:i*10+10]

            with pdf.table(text_align=Align.C, col_widths=100, line_height=10, padding=2) as table:
                for data_row in TABLE_DATA:
                    row = table.row()
                    for datum in data_row:
                        row.cell(datum, padding=(8, 5, 8, 5))

            pdf.set_left_margin(30)

    ################# page (Tier 4 Suggestions sorted by Inbreeding Coefficient) #################
        sorted_tier4_sugs = sortByCoi(tier4_sugs, genType)
        for i in range(math.ceil(len(sorted_tier4_sugs) / 10)):
            page_label = "TOP STALLIONS BY INBREEDING COEFFICIENT"
            if i != 0:
                page_label += "(CONTINUED)"
            pdf.add_page()
            pdf.set_line_width(2)
            pdf.set_fill_color(r=255, g=255, b=255)
            pdf.rect(x=50, y=80, w=280, h=90, style="D")
            pdf.rect(x=450, y=80, w=100, h=70, style="D")

            pdf.ln()
            pdf.ln()
            pdf.set_font('Times', '', 25)
            pdf.set_text_color(0, 0, 0)
            pdf.cell(lmargin+10)
            pdf.cell(w=0, h=30, text="Tier 4", new_x=XPos.LMARGIN, new_y=YPos.TOP)

            pdf.set_font('Times', '', 60)
            pdf.cell(420)
            pdf.cell(w=100, h=40, text=f"{len(tier4_sugs)}", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.set_font('Times', '', 18)
            pdf.cell(lmargin+10)
            pdf.multi_cell(w=260, h=20, text=f"{tier4_filter_label}", new_x=XPos.LMARGIN, new_y=YPos.TOP)

            pdf.set_font('Times', '', 10)
            pdf.cell(420)
            pdf.cell(w=100, h=25, text="MATCHES FOUND", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.ln(100)

            pdf.set_font('Times', '', 18)
            pdf.cell(lmargin)
            pdf.cell(w=0, h=0, text=page_label, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.ln(100)
            pdf.set_left_margin(70)
            pdf.set_line_width(0.5)
            pdf.set_font('Times', '', 10)
            TABLE_HEADER_DATA = [
                ["Stallion", "1D Rate", "Variant", "Equi-Source Score", "Inbreeding Coefficient of foal"]
            ]
            TABLE_DATA = TABLE_HEADER_DATA + sorted_tier4_sugs[i*10:i*10+10]

            with pdf.table(text_align=Align.C, col_widths=100, line_height=10, padding=2) as table:
                for data_row in TABLE_DATA:
                    row = table.row()
                    for datum in data_row:
                        row.cell(datum, padding=(8, 5, 8, 5))

            pdf.set_left_margin(30)

    ################# page (Top Ancestors) #################
    if len(anc_top_data) != 0:
        sorted_anc_top_data = sortByIndex(anc_top_data, 3)
        for i in range(math.ceil(len(sorted_anc_top_data) / 10)):
            page_label = "TOP ANCESTORS"
            if i != 0:
                page_label += "(CONTINUED)"
            pdf.add_page()
            pdf.set_line_width(2)
            pdf.set_fill_color(r=255, g=255, b=255)
            pdf.rect(x=50, y=80, w=280, h=70, style="D")
            pdf.rect(x=450, y=80, w=100, h=70, style="D")

            pdf.ln()
            pdf.ln()
            pdf.set_font('Times', '', 25)
            pdf.set_text_color(0, 0, 0)
            pdf.cell(lmargin+10)
            pdf.cell(w=280, h=30, text=sheetName, new_x=XPos.LMARGIN, new_y=YPos.TOP)

            pdf.set_font('Times', '', 60)
            pdf.set_text_color(grade_color[0], grade_color[1], grade_color[2])
            pdf.cell(420)
            pdf.cell(w=100, h=40, text=letter_grade, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.set_font('Times', '', 18)
            pdf.set_text_color(0, 0, 0)
            pdf.cell(lmargin+10)
            pdf.cell(w=280, h=10, text=f"{pedigree_dict['birth']} {pedigree_dict['sex']}", new_x=XPos.LMARGIN, new_y=YPos.TOP)

            pdf.set_font('Times', '', 10)
            pdf.set_text_color(0, 0, 0)
            pdf.cell(420)
            pdf.cell(w=100, h=25, text=f"VARIANT = {v_sum}", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.ln(100)

            pdf.set_font('Times', '', 18)
            pdf.cell(lmargin)
            pdf.cell(w=0, h=0, text=page_label, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.ln(100)
            pdf.set_left_margin(70)
            pdf.set_line_width(0.5)
            pdf.set_font('Times', '', 10)
            TABLE_HEADER_DATA = [
                ["Top Ancestor Stallions", "Total Frequency", "Position Diversity", "Position Flexibility Score", "Inbreeding Coefficient"]
            ]
            TABLE_DATA = TABLE_HEADER_DATA + sorted_anc_top_data[i*10:i*10+10]

            with pdf.table(text_align=Align.C, col_widths=100, line_height=10, padding=2) as table:
                for data_row in TABLE_DATA:
                    row = table.row()
                    for datum in data_row:
                        row.cell(datum, padding=(8, 5, 8, 5))

            pdf.set_left_margin(30)

    ################# page (Ancestor Position And Frequency) #################
    if len(anc_pedigree_data) != 0:
        pdf.add_page()
        pdf.set_line_width(2)
        pdf.set_fill_color(r=255, g=255, b=255)
        pdf.rect(x=50, y=80, w=280, h=70, style="D")
        pdf.rect(x=450, y=80, w=100, h=70, style="D")

        pdf.ln()
        pdf.ln()
        pdf.set_font('Times', '', 25)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(lmargin+10)
        pdf.cell(w=280, h=30, text=sheetName, new_x=XPos.LMARGIN, new_y=YPos.TOP)

        pdf.set_font('Times', '', 60)
        pdf.set_text_color(grade_color[0], grade_color[1], grade_color[2])
        pdf.cell(420)
        pdf.cell(w=100, h=40, text=letter_grade, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        pdf.set_font('Times', '', 18)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(lmargin+10)
        pdf.cell(w=280, h=10, text=f"{pedigree_dict['birth']} {pedigree_dict['sex']}", new_x=XPos.LMARGIN, new_y=YPos.TOP)

        pdf.set_font('Times', '', 10)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(420)
        pdf.cell(w=100, h=25, text=f"VARIANT = {v_sum}", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        pdf.ln(100)

        pdf.set_font('Times', '', 18)
        pdf.cell(lmargin)
        pdf.cell(w=0, h=0, text="ANCESTOR POSITION AND FREQUENCY", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        pdf.set_line_width(0.5)
        pdf.rect(x=50, y=450, w=120, h=25, style="D")
        pdf.rect(x=50, y=610, w=120, h=25, style="D")

        pdf.rect(x=180, y=410, w=120, h=25, style="D")
        pdf.rect(x=180, y=490, w=120, h=25, style="D")
        pdf.rect(x=180, y=570, w=120, h=25, style="D")
        pdf.rect(x=180, y=650, w=120, h=25, style="D")

        pdf.rect(x=310, y=390, w=120, h=25, style="D")
        pdf.rect(x=310, y=430, w=120, h=25, style="D")
        pdf.rect(x=310, y=470, w=120, h=25, style="D")
        pdf.rect(x=310, y=510, w=120, h=25, style="D")
        pdf.rect(x=310, y=550, w=120, h=25, style="D")
        pdf.rect(x=310, y=590, w=120, h=25, style="D")
        pdf.rect(x=310, y=630, w=120, h=25, style="D")
        pdf.rect(x=310, y=670, w=120, h=25, style="D")

        pdf.rect(x=440, y=390, w=120, h=25, style="D")
        pdf.rect(x=440, y=430, w=120, h=25, style="D")
        pdf.rect(x=440, y=470, w=120, h=25, style="D")
        pdf.rect(x=440, y=510, w=120, h=25, style="D")
        pdf.rect(x=440, y=550, w=120, h=25, style="D")
        pdf.rect(x=440, y=590, w=120, h=25, style="D")
        pdf.rect(x=440, y=630, w=120, h=25, style="D")
        pdf.rect(x=440, y=670, w=120, h=25, style="D")

        pdf.ln(150)
        pdf.set_font('Times', '', 10)
        pdf.cell(280)
        pdf.cell(w=120, h=0, text=f"{anc_pedigree_data[5]} Top Ancestors", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP) # 1/3
        pdf.cell(410)
        pdf.cell(w=120, h=0, text=f"{anc_pedigree_data[9]} Top Ancestors", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP) # 1/4
        pdf.ln(20)
        pdf.cell(150)
        pdf.cell(w=120, h=0, text=f"{anc_pedigree_data[4]} Top Ancestors", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP) # 1/2
        pdf.ln(20)
        pdf.cell(410)
        pdf.cell(w=120, h=0, text=f"{anc_pedigree_data[10]} Top Ancestors", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP) # 2/4
        pdf.ln(20)
        pdf.cell(20)
        pdf.cell(w=120, h=0, text=f"{anc_pedigree_data[2]} Top Ancestors", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP) # 1/1
        pdf.ln(20)
        pdf.cell(280)
        pdf.cell(w=120, h=0, text=f"{anc_pedigree_data[7]} Top Ancestors", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP) # 3/3
        pdf.cell(410)
        pdf.cell(w=120, h=0, text=f"{anc_pedigree_data[11]} Top Ancestors", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP) # 3/4
        pdf.ln(20)
        pdf.ln(20)
        pdf.cell(410)
        pdf.cell(w=120, h=0, text=f"{anc_pedigree_data[12]} Top Ancestors", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT) # 4/4
        pdf.ln(40)
        pdf.cell(280)
        pdf.cell(w=120, h=0, text=f"{anc_pedigree_data[6]} Top Ancestors", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP) # 5/3
        pdf.cell(410)
        pdf.cell(w=120, h=0, text=f"{anc_pedigree_data[13]} Top Ancestors", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP) # 5/4
        pdf.ln(20)
        pdf.cell(150)
        pdf.cell(w=120, h=0, text=f"{anc_pedigree_data[3]} Top Ancestors", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP) # 3/2
        pdf.ln(20)
        pdf.cell(410)
        pdf.cell(w=120, h=0, text=f"{anc_pedigree_data[14]} Top Ancestors", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP) # 6/4
        pdf.ln(20)
        pdf.ln(20)
        pdf.cell(280)
        pdf.cell(w=120, h=0, text=f"{anc_pedigree_data[8]} Top Ancestors", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP) # 7/3
        pdf.cell(410)
        pdf.cell(w=120, h=0, text=f"{anc_pedigree_data[15]} Top Ancestors", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP) # 7/4
        pdf.ln(20)
        pdf.ln(20)
        pdf.cell(410)
        pdf.cell(w=120, h=0, text=f"{anc_pedigree_data[16]} Top Ancestors", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP) # 8/4
    
    ################# page (Frequency of Top Ancestors by Stallion) #################
    if len(stallion_data) != 0:
        pdf.add_page()
        pdf.set_line_width(2)
        pdf.set_fill_color(r=255, g=255, b=255)
        pdf.rect(x=50, y=80, w=280, h=70, style="D")
        pdf.rect(x=450, y=80, w=100, h=70, style="D")

        pdf.ln()
        pdf.ln()
        pdf.set_font('Times', '', 25)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(lmargin+10)
        pdf.cell(w=280, h=30, text=sheetName, new_x=XPos.LMARGIN, new_y=YPos.TOP)

        pdf.set_font('Times', '', 60)
        pdf.set_text_color(grade_color[0], grade_color[1], grade_color[2])
        pdf.cell(420)
        pdf.cell(w=100, h=40, text=letter_grade, align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        pdf.set_font('Times', '', 18)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(lmargin+10)
        pdf.cell(w=280, h=10, text=f"{pedigree_dict['birth']} {pedigree_dict['sex']}", new_x=XPos.LMARGIN, new_y=YPos.TOP)

        pdf.set_font('Times', '', 10)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(420)
        pdf.cell(w=100, h=25, text=f"VARIANT = {v_sum}", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        pdf.ln(100)

        pdf.set_font('Times', '', 18)
        pdf.cell(lmargin)
        pdf.cell(w=0, h=0, text="FREQUENCY OF TOP ANCESTORS BY STALLION", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        pdf.ln(50)
        pdf.set_left_margin(400)
        pdf.set_line_width(0.5)
        pdf.set_font('Times', '', 10)
        TABLE_HEADER_DATA = [
            ["Stallions", "Frequency"]
        ]
        TABLE_DATA = TABLE_HEADER_DATA + sortByIndex(stallion_data, 1)

        with pdf.table(text_align=Align.C, col_widths=100, line_height=10, padding=2) as table:
            for data_row in TABLE_DATA:
                row = table.row()
                for datum in data_row:
                    row.cell(datum, padding=(8, 5, 8, 5))

        pdf.set_left_margin(30)

    pdf.output(f"{sheetName}.pdf")
    return {"status": 0, "msg": "Success"}

def load_spreadsheet_data(wsheetId, msheetId):
    worksheet = getGoogleSheetService().spreadsheets()
    sheet_names = []
    try:
        wsheet_metadata = worksheet.get(spreadsheetId=wsheetId).execute()
        for sheet in wsheet_metadata['sheets']:
            sheet_names.append(sheet['properties']['title'])
        try:
            worksheet.get(spreadsheetId=msheetId).execute()
            return {"status": 0, "msg": "Success", "data": sheet_names}
        except:
            return {"status": -1, "msg": "The mastersheet ID is not valid."}
    except:
        return {"status": -1, "msg": "The worksheet ID is not valid."}
    
create_pdf(wsheetId="1ta2t7cQarx6dzacKHLevtYYOsnV_gei9N2z7lHvlhLs", sheetName="Master Heidi Flys", msheetId="18wZ_UlyQKmhzygdb8nk8I6xAyIPvxJm3Ofh58d1NKZs", genType=0)