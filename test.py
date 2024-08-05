from fpdf import FPDF, XPos, YPos, Align
import math
import requests
from requests.adapters import HTTPAdapter
from bs4 import BeautifulSoup
from constants import *

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
    ##############################################################
    ###################                        ###################
    ################### PDF Generation Process ###################
    ###################                        ###################
    ##############################################################
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
    grade_info = getGradeInfo("A", "B-", "B", "A")
    letter_grade = grade_info["letter"]
    grade_color = grade_info["color_info"]
    pdf.add_page()
    pdf.set_line_width(2)
    pdf.set_fill_color(r=255, g=255, b=255)
    pdf.rect(x=50, y=100, w=240, h=70, style="D")
    pdf.rect(x=390, y=100, w=90, h=70, style="D")
    pdf.image('assets/images/tempbar.png', 650, 120, 200, 60)
    pdf.image('assets/images/teardrop.png', grade_info["tempbar_pos"], 75, 25)
    
    pdf.ln()
    
    pdf.set_font('Times', 'B', 6)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(grade_info["tempbar_pos"]-27)
    pdf.cell(w=22, h=20, text="108.24", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.TOP)
    
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
    pdf.cell(w=280, h=10, text=f"2017 Mare", new_x=XPos.LMARGIN, new_y=YPos.TOP)

    pdf.set_font('Times', '', 9)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(360)
    pdf.cell(w=90, h=25, text=f"VARIANT = 108.24", align=Align.C, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    
    pdf.set_line_width(0.5)
    pdf.ln(92)

    pdf.output(f"{wsheetName}.pdf")
    return {"status": MSG_SUCCESS, "msg": "Success"}
    
create_pdf(wsheetId="1h-tZdm0-UJnC09j8dYidTND1FCWRGDxkBMCmHzr1bYM", wsheetName="Mistys Money N Fame", msheetId="1g5kX6F34q2HFn4aqfXb5tkjBM_qTSy4fHUakxz6qJj0", genType=0)