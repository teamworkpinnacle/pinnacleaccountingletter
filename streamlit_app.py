from st_pages import hide_pages
from time import sleep
import streamlit as st

if 'user_name' not in st.session_state:
    st.session_state['user_name'] = ''

if 'password' not in st.session_state:
    st.session_state['password'] = ''




def log_in():
    st.session_state["logged_in"] = True
    hide_pages([])
    st.success("Logged in!")
    sleep(0.5)
    st.switch_page("pages/page1.py")


def log_out():
    st.session_state["logged_in"] = False
    st.success("Logged out!")
    sleep(0.5)


if not st.session_state.get("logged_in", False):
    hide_pages(["page1", "page2", "page3"])
    username = st.text_input("Username", key="username")
    password = st.text_input("Password", key="password", type="password")

    if username == "test" and password == "test":
        st.session_state["logged_in"] = True
        hide_pages([])
        st.success("Logged in!")
        sleep(0.5)
        st.switch_page("pages/page1.py")

else:
    st.write("Logged in!")
    st.button("log out", on_click=log_out)