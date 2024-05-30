import streamlit as st
from docxtpl import DocxTemplate
import io

# function to format document
def render_docx(data):
    doc = DocxTemplate(report_dict[st.session_state['report_sample']])
    doc.render(data)
    return doc

# set page configuration
st.set_page_config(
    page_title="Sample Report Generator",
    layout="wide")

# initialize variables in session state
if 'Context' not in st.session_state:
	st.session_state['Context'] = {'Title_of_Report' : '[Report Title]',
	'Faculty_Advisor': '[Faculty Advisor]',
	'Org_Advisor' : '[Organization Advisor]',
    'Org_Name' : '[Organization Name]',
    'Date' : '[xx Month 20xx]',
    'Student_Name' : "[Student Name]",
    'Student_Number' : '[Student Number]',
    'Year' : '[Year]',
    'Program' : '[Program]'}

if 'report_sample' not in st.session_state:
	st.session_state['report_sample'] = "Choose Report"

report_dict = {'DAP Report':'DAP Sample report.docx',
				'SIP Report':'SIP Sample report.docx'}

st.markdown("Enter your information here:")

# which report
report_list = ['DAP Report', 'SIP Report']
st.session_state['report_sample'] = st.selectbox('Which program are you writing this report for?', report_list, placeholder="Choose Report")

st.session_state['Context']['Title_of_Report'] = st.text_input('What is the Title of your report?', key='report_name',placeholder='[Report Title]')

st.session_state['Context']['Student_Name'] = st.text_input('What is your full name?', key='student_name',placeholder='[Student Name]')

st.session_state['Context']['Student_Number'] = st.text_input('What is your student id number?', key='student_no',placeholder='[Student Number]')

st.session_state['Context']['Program'] = st.selectbox('Which program are you writing this report for?', ["UG", "PG"], placeholder="[Program]")

st.session_state['Context']['Year'] = st.text_input('Which year did you join in?', key='year',placeholder='[Year]')

st.session_state['Context']['Date'] = st.text_input('Which is the date today?', key='date',placeholder='eg: 29 May 2024')

st.session_state['Context']['Faculty_Advisor'] = st.text_input('Full name of Faculty Advisor', key='faculty',placeholder='eg: Prof. Viraj Shah')

st.session_state['Context']['Org_Advisor'] = st.text_input('Full name of Organization Advisor', key='org_adv',placeholder='eg: Mr. XYZ')

st.session_state['Context']['Org_Name'] = st.text_input('Full name of Organization', key='org_name',placeholder='eg: ABC')


if st.button("Submit Form"):
	if st.session_state['report_sample']:
		if st.session_state['Context']['Title_of_Report']:
			if st.session_state['Context']['Student_Name']:
				if st.session_state['Context']['Student_Number']:
					if st.session_state['Context']['Program']:
						if st.session_state['Context']['Year']:
							if st.session_state['Context']['Date']:
								if st.session_state['Context']['Faculty_Advisor']:
									if st.session_state['Context']['Org_Advisor']:
										if st.session_state['Context']['Org_Name']:
											doc = render_docx(st.session_state['Context'])
											bio = io.BytesIO()
											doc.save(bio)

											st.download_button(label="Download Template Report",
												data=bio.getvalue(),
												file_name='report.docx',
												mime='docx')

st.markdown("*if you do not get a button to download the template after submitting the form, please check if you have filled information in all the above mentioned fields.")
