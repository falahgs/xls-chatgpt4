import streamlit as st
import pandas as pd
from langchain.agents import create_csv_agent
from langchain.llms import OpenAI
import time
from streamlit_option_menu import option_menu
# Set Streamlit page configuration
st.set_page_config(page_title='XLS ChatGPT4', 
                   page_icon=":memo:", 
                   layout='wide', 
                   initial_sidebar_state='collapsed')

# Set CSS properties for HTML components
st.markdown("""
<style>
body {
    color: #fff;
    background-color: #4f8bf9;
}
h1, h2 {
    color: #ffdd00;
}
</style>
    """, unsafe_allow_html=True)

hide_style='''
    <style>
    #MainMenu {visibility:hidden;}
    footer {visibility:hidden;}
    .css-hi6a2p {padding-top: 0rem;}
    head {visibility:hidden;}
    </style>
'''
st.markdown("""
<h1 style='text-align: center; color: #ffdd00;'>XLS Office Documents Analysis with ChatGPT4 NLP Model</h1>
    """, unsafe_allow_html=True)
#st.title('XLS Office Documents Analysis with ChatGPT4 NLP Model')
# 1. as sidebar menu
with st.sidebar:
    selected = option_menu("Main Menu", ["Home", 'Help'], 
        icons=['house', 'gear'], menu_icon="cast", default_index=0)
    selected
    hide_style='''
    <style>
    #MainMenu {visibility:hidden;}
    footer {visibility:hidden;}
    .css-hi6a2p {padding-top: 0rem;}
    head {visibility:hidden;}
    </style>
'''
hide_streamlit_style = """
                <style>
                div[data-testid="stToolbar"] {
                visibility: hidden;
                height: 0%;
                position: fixed;
                }
                div[data-testid="stDecoration"] {
                visibility: hidden;
                height: 0%;
                position: fixed;
                }
                div[data-testid="stStatusWidget"] {
                visibility: hidden;
                height: 0%;
                position: fixed;
                }
                #MainMenu {
                visibility: hidden;
                height: 0%;
                }
                header {
                visibility: hidden;
                height: 0%;
                }
                footer {
                visibility: hidden;
                height: 0%;
                }
                </style>
                """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)
st.markdown("""
<style>
.footer {
  position: fixed;
  left: 0;
  bottom: 0;
  width: 100%;
  background-color: #f8f9fa;
  color: black;
  text-align: center;
  padding: 10px;
}
</style>

<div class="footer">
<p>Copyright &copy; 2023 AI-Books. All Rights Reserved.</p>
</div>
""", unsafe_allow_html=True)

if selected=="Help":

    st.markdown(hide_style,unsafe_allow_html=True)
   # st.title("Help")
    # Import required libraries
    import streamlit as st

    # Set Streamlit page configuration
    #st.set_page_config(page_title='Help - XLS Office Documents Analysis with ChatGPT4 NLP Model', 
                      # page_icon=":memo:", 
                      # layout='wide', 
                      # initial_sidebar_state='collapsed')

    # Set CSS properties for HTML components
    st.markdown("""
    <style>
    body {
        color: #fff;
        background-color: #4f8bf9;
    }
    h1, h2 {
        color: #ffdd00;
    }
    </style>
        """, unsafe_allow_html=True)

    # Use HTML in markdown to center align the title
   # st.markdown("""
   # <h1 style='text-align: center; color: #ffdd00;'>Help Document for XLS Office Documents Analysis with ChatGPT4 NLP Model</h1>
      #  """, unsafe_allow_html=True)

    # Display authorship details
    st.markdown("""
    ## Developed by Falah.G.Salieh
    * AI Developer
    * Specialized in Natural Language Processing
    """, unsafe_allow_html=True)

    # Display help content
    st.markdown("""
    ## Getting Started

    This project aims to provide a tool for data analysis on any office document. You can simply upload an Excel spreadsheet and ask any questions related to your data. Our system, powered by the ChatGPT-4 Natural Language Processing Model, will parse your data and respond to your questions. 

    Here's how to use this app:

    1. Upload your XLS file: Use the file uploader to select and upload your XLS file.
    2. Enter your question: Ask a question related to your data in the question field. You can choose a question example from the dropdown list or type in a question on your own.
    3. Get your answer: Click 'Run' and the system will analyze your data and provide a response to your question.

    Please note that this system uses a natural language processing model which may not always provide a perfect response, but it strives to deliver the best possible answer based on the data given.

    ### Uploading Your XLS File

    To upload your XLS file, click on the 'Upload XLS' button and select your file from your system. The uploaded file will be converted to CSV format to allow for data analysis.

    ### Using the ChatGPT-4 NLP Model

    The ChatGPT-4 model is a powerful language model that can answer questions based on the data provided. Enter a question related to your data in the question field. The model will parse your question, analyze the data, and generate a response.

    ### Understanding the Results

    When you receive the results, they will be displayed in the 'Response' section. The response is generated by the ChatGPT-4 model based on your question and the data from the uploaded file.

    ### Troubleshooting

    If you encounter any problems, please ensure that your XLS file is valid and that your question is clear and related to the data in your file.

    ## Contact

    If you have any questions, feel free to contact the developer.
    """, unsafe_allow_html=True)
    


#-----------------




if selected=="Home":
    #st.title("home")

    def load_data(file):
        df = pd.read_excel(file, engine='openpyxl')
        df.to_csv('data.csv', index=False)  # Convert XLS to CSV
        return 'data.csv'

    def initialize_agent(file, openai_api_key):
        agent = create_csv_agent(OpenAI(temperature=0, openai_api_key=openai_api_key), file, verbose=True)
        return agent

    uploaded_file = st.file_uploader("Upload XLS", type=['xlsx'])
    st.markdown(hide_style, unsafe_allow_html=True)

    openai_api_key = st.sidebar.text_input('OpenAI API Key', type="password")

    # Pre-defined question examples
    question_examples = [
        "how many rows are there?",
        "how many people are female?",
        "how many people have stayed more than 3 years in the city?",
        "how many people have stayed more than 3 years in the city and are female?",
        "Are there more males or females?",
        "What are the column names?",
        "What is the average age?",
        "Which country appears the most and how many times does it appear?",
        "What is the ratio of males to females?"
        # Add more examples as needed
    ]

    # Dropdown select box for question examples
    selected_example = st.selectbox('Choose a question example:', question_examples)

    # Pre-populate the question field with the selected example
    question = st.text_input('Enter your question:', value=selected_example)

    if not openai_api_key or not openai_api_key.startswith('sk-'):
        st.warning('Please enter your OpenAI API key!', icon='⚠️')
    else:
        if uploaded_file is not None:
            # Create a progress bar
            #progress_bar = st.progress(0)
            #progress_bar.progress(25)  # Start the progress at 25%
            
            csv_file = load_data(uploaded_file)  # Now the uploaded file is an XLS file
            #progress_bar.progress(50)  # Update the progress to 50%
            
            agent = initialize_agent(csv_file, openai_api_key)
            #progress_bar.progress(100)  # Complete the progress bar
            
            if question:
                response = agent.run(question)
                with st.spinner('Wait for it...'):
                    time.sleep(5)
                st.success('Done!')
                #st.markdown(f'**Response:** {response}')
                st.markdown(f'<div style="color: red; font-size: 24px; text-align: center;">The Answer is:{response}</div>',unsafe_allow_html=True)

