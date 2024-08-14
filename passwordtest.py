
import streamlit as st
from streamlit_gsheets import GSheetsConnection
login = None
conn = st.connection("gsheets", type=GSheetsConnection)
if 'user_name' not in st.session_state:
    st.session_state['user_name'] = ''

if 'password' not in st.session_state:
    st.session_state['password'] = ''

# if login == True:
#     st.write("Mehmeh")

# else:
    # user_name = st.text_input('Username', placeholder = 'username@email.com')
    # user_password =  st.text_input('Password', placeholder = '12345678',type="password")
user_name = "Hans"
user_password =  "Pinnacle8539!"
login = st.button("Login", type="primary")
user_crediantials = conn.read(worksheet="Crediantials")
user_crediantials_list = user_crediantials["username"].to_list()
user_crediantials = user_crediantials.set_index("username")
print(user_crediantials)

if  user_name in user_crediantials_list:
    print("Correct Username")
    if user_crediantials["password"][user_name] == user_password:
        # login = True
        print("Correct Password")


    else:
        print("WRONG PASSWORD")


# print(user_crediantials)
# gs_user_db = user_crediantials.T.to_dict()
# print(gs_user_db)

# print(user_crediantials["username"].to_list())
st.session_state['user_name'] = user_name
st.session_state['password'] = user_password

