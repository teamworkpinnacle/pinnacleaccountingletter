from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.enum.style import WD_STYLE_TYPE
import numpy as np
from pathlib import Path
import os
import io
import time
import streamlit as st
from streamlit_gsheets import GSheetsConnection

# Page Configuration
st.set_page_config(
    page_title='Pinnacle Accounting Cover Letter2',
    page_icon=':white_check_mark:',
    layout='wide'
)

downloads_path = str(Path.home() / "Downloads")
doc_template_path = "data/coverpage.docx"

# Create a connection object.
conn = st.connection("gsheets", type=GSheetsConnection)

# Initializing session state variables
if 'user_name' not in st.session_state:
    st.session_state['user_name'] = ''

if 'password' not in st.session_state:
    st.session_state['password'] = ''

if "word_doc_button_clicked" not in st.session_state:
    st.session_state.word_doc_button_clicked = False

if 'doc' not in st.session_state:
    st.session_state['doc'] = None

if 'docname' not in st.session_state:
    st.session_state['docname'] = ''


def password_checker(user_name, user_password):
    user_credentials = conn.read(worksheet="Crediantials")
    user_credentials_list = user_credentials["username"].to_list()
    user_credentials = user_credentials.set_index("username")

    if user_name in user_credentials_list:
        if user_credentials["password"][user_name] == user_password:
            return True
        return "Password is incorrect"
    return "Username is incorrect"


def add_new_doc(df, doc):
    for x in range(len(df)):
        if df["Type"][x] == "new paragraph":
            paragraph = doc.add_paragraph()

        elif df["Type"][x] == "paragraph":
            font = paragraph.add_run(f'''{df["Content"][x]}''').font
            font.name = df["fontname"][x]
            font.size = Pt(df["fontsize"][x])
            font.bold = df["fontbold"][x]
            font.underline = df["fontunderline"][x]
            font.highlight_color = df["fonthighlight_color"][x]

        elif df["Type"][x] == "table":
            table = doc.add_table(rows=0, cols=2)
            table.style = doc.styles['Table Grid']
            table.autofit = True
            table.allow_autofit = True

        elif df["Type"][x] == "add on to table":
            row_cells = table.add_row().cells
            tableparagraph = row_cells[0].paragraphs[0]
            run = tableparagraph.add_run(f"{df['Content'][x]}")
            run.underline = df["fontunderline"][x]
            run.font.size = Pt(df["fontsize"][x])
            run.font.name = df["fontname"][x]
            run.font.highlight_color = df["fonthighlight_color"][x]
            tableparagraph = row_cells[1].paragraphs[0]
            run = tableparagraph.add_run(f"{df['content2'][x]}")
            run.underline = df["fontunderline"][x]
            run.font.size = Pt(df["fontsize"][x])
            run.font.name = df["fontname"][x]
            run.font.highlight_color = df["fonthighlight_color"][x]
    return doc


def sel_callback():
    for i in range(1, 26):
        st.session_state[f'col{i}'] = st.session_state.sel


st.header(":ballot_box_with_check: Select the content you would like to include in the Accounting Letter")
cola, colb, colc = st.columns(3)

# Checkbox section
with cola:
    Property_plant_equipment = st.checkbox("Property plant equipment", value=False, key="col1")
    Investment_properties = st.checkbox("Investment properties", value=False, key="col2")
    Fair_value = st.checkbox("Fair value", value=False, key="col3")
    Investment_in_joint_venture = st.checkbox("Investment in joint venture", value=False, key="col4")
    Investment_in_associates = st.checkbox("Investment in associates", value=False, key="col5")
    Investment_in_subsidiaries = st.checkbox("Investment in subsidiaries", value=False, key="col6")
    Intangible_assets = st.checkbox("Intangible assets", value=False, key="col7")
    Goodwill = st.checkbox("Goodwill", value=False, key="col8")
    st.write("")
    st.checkbox('Select All', key='sel', on_change=sel_callback)

with colb:
    Right_of_use = st.checkbox("Right of use", value=False, key="col9")
    Inventory = st.checkbox("Inventory", value=False, key="col10")
    Cash_and_bank_balances = st.checkbox("Cash and bank balances", value=False, key="col11")
    Trade_and_other_receivables = st.checkbox("Trade and other receivables", value=False, key="col12")
    Trade_and_other_payables = st.checkbox("Trade and other payables", value=False, key="col13")
    Borrowings = st.checkbox("Borrowings", value=False, key="col14")
    Amount_due = st.checkbox("Amount due", value=False, key="col15")
    Revenue_recognition = st.checkbox("Revenue recognition", value=False, key="col16")
    Gross_profit_margin = st.checkbox("Gross profit margin", value=False, key="col17")

with colc:
    Voluntary_cpf_contributions = st.checkbox("Voluntary cpf contributions", value=False, key="col18")
    Government_grant = st.checkbox("Government grant", value=False, key="col19")
    Foreign_exchange = st.checkbox("Foreign exchange", value=False, key="col20")
    Small_value_assets = st.checkbox("Small value assets", value=False, key="col21")
    Presentation_currency = st.checkbox("Presentation currency", value=False, key="col22")
    Expenses_recognition = st.checkbox("Expenses recognition", value=False, key="col23")
    Gst_registration = st.checkbox("Gst registration", value=False, key="col24")
    Representations_from_the_company = st.checkbox("Representations from the company", value=False, key="col25")

mehmehlist = [Property_plant_equipment, Investment_properties, Fair_value, Investment_in_joint_venture,
              Investment_in_associates, Investment_in_subsidiaries, Intangible_assets, Goodwill, Right_of_use,
              Inventory, Cash_and_bank_balances, Trade_and_other_receivables,
              Trade_and_other_payables, Borrowings, Amount_due, Revenue_recognition, Gross_profit_margin,
              Voluntary_cpf_contributions, Government_grant, Foreign_exchange, Small_value_assets, Presentation_currency,
              Expenses_recognition, Gst_registration, Representations_from_the_company]

list_of_sheets = ["PROPERTY, PLANT AND EQUIPMENT", "INVESTMENT PROPERTIES", "FAIR VALUE",
                  "INVESTMENT IN JOINT VENTURE", "INVESTMENT IN ASSOCIATES", "INVESTMENT IN SUBSIDIARIES",
                  "INTANGIBLE ASSETS", "GOODWILL", "RIGHT OF USE ASSETS & LEASE LIABILITIES", "INVENTORIES",
                  "CASH AND BANK BALANCES", "TRADE AND OTHER RECEIVABLES", "TRADE AND OTHER PAYABLES", "BORROWINGS",
                  "AMOUNT DUE FROM/ TO SHAREHOLDERS/ DIRECTORS", "REVENUE RECOGNITION, CONTRACT ASSETS AND CONTRACT LIABILITIES",
                  "GROSS PROFIT MARGIN", "VOLUNTARY CPF CONTRIBUTIONS", "GOVERNMENT GRANT – CAPITAL", "FOREIGN EXCHANGE",
                  "SMALL VALUE ASSETS", "PRESENTATION CURRENCY", "EXPENSES RECOGNITION", "GST REGISTRATION", "REPRESENATATIONS FROM THE COMPANY"]


def checker():
    doc = Document(doc_template_path)
    for x in range(len(mehmehlist)):
        if mehmehlist[x] == True:
            df = conn.read(worksheet=list_of_sheets[x])
            df = df.replace(np.nan, '', regex=True)
            df["fonthighlight_color"] = df["fonthighlight_color"].replace("None", None, regex=True)
            df["fonthighlight_color"] = df["fonthighlight_color"].replace("turquoise", WD_COLOR_INDEX.TURQUOISE, regex=True)
            df["fonthighlight_color"] = df["fonthighlight_color"].replace("yellow", WD_COLOR_INDEX.YELLOW, regex=True)
            doc = add_new_doc(df, doc)
            if sum(mehmehlist) > 19:
                time.sleep(4.1)
    return doc


def Word_Document_Created():
    st.session_state.word_doc_button_clicked = True


username = st.text_input("Username", key="username")
password = st.text_input("Password", key="password", type="password")
docname = st.text_input(label="Document Name", placeholder="Insert Word Document File Name Here Before Clicking Create")

if st.button("Create Word Document", on_click=Word_Document_Created) or st.session_state.word_doc_button_clicked:
    if st.session_state['doc'] is None or st.session_state['docname'] != docname:
        check = password_checker(username, password)
        if check == True:
            st.subheader("After loading is done, press the button below to download. File will be in your download folder")
            st.session_state['doc'] = checker()
            st.session_state['docname'] = docname
        else:
            st.error(check)
            st.session_state.word_doc_button_clicked = False
            st.session_state['doc'] = None

    if st.session_state['doc']:
        bio = io.BytesIO()
        st.session_state['doc'].save(bio)
        st.download_button(
            label="Click here to download",
            data=bio.getvalue(),
            file_name=f"{st.session_state['docname']}.docx",
            mime="docx"
        )


