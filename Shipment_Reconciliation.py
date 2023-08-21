import streamlit as st 
import pandas as pd
from shipment_reco_charts import reconcile
import os
import streamlit_authenticator as stauth
import pickle 
from pathlib import Path 
import yaml
from PIL import Image
import time
import plotly.graph_objects as go
import base64
from page_config import page_setup
from login_page import login_status


st.set_page_config(layout="wide",initial_sidebar_state ="collapsed")

page_setup()

state = st.session_state

authentication_status = login_status()

# st.markdown('''
#     <style>
#     .css-9s5bis.edgvbvh3 {
#     display: none;
#     }
#     </style>
#     ''', unsafe_allow_html=True)

# with open('config.yaml') as file:
#     config = yaml.safe_load(file)

# authenticator = stauth.Authenticate(
#     config['credentials'],
#     config['cookie']['name'],
#     config['cookie']['key'],
#     config['cookie']['expiry_days'],
#     config['preauthorized']
# )
# state.authenticator = authenticator

# #placeholder = st.empty()

# #authenticator = stauth.Authenticate(names, usernames, hashed_passwords, "ship_recon", "admin")
# #with placeholder.container():
# space, login, space = st.columns([1,3,1])
# with login:
#     name, authentication_status, username = authenticator.login('Login', 'main')
# state.authentication_status

if authentication_status == False:
    space, login, space = st.columns([1,3,1])
    with login:
        st.error("Username/Password is incorrect")



if authentication_status:
    #placeholder.empty()
    #authenticator.logout('Logout', 'sidebar')

    time.sleep(0.1)
    def landing_page():
        st.markdown('''
        <style>
        .css-9s5bis.edgvbvh3 {
        display: block;
        }
        </style>
        ''', unsafe_allow_html=True)
        #with title:
        # emp,title,emp = st.columns([2,2,2])
        # with title:
        if 'submit' not in state:
            state.submit= False
        if 'response' not in state:
            state.response = []
        st.markdown("<h2 style='text-align: center; padding:0'>Shipment Reconciliation</h2>", unsafe_allow_html=True)
        #st.write('###')
        shipment_instructions, warehouse_reports, inventory_ledger, submit = file_upload_form()
        #print(warehouse_reports)
        # try:
        if submit:
            state.submit = True
            #print(warehouse_reports)
            #print(submit)
                #print(shipment_instructions_df)
            with st.spinner('Please wait'):
                try:
                    delete_temp()
                except:
                    print()
                if inventory_ledger is not None:
                    inventory_ledger_df = pd.read_csv(inventory_ledger)
                units_booked, excess_units_received, short_units_received, units_received, matching_sku, mismatching_sku = reconcile(shipment_instructions, warehouse_reports, inventory_ledger_df)
                state.response = [units_booked, excess_units_received, short_units_received, units_received, matching_sku, mismatching_sku]
                
                bar_data = [['Units Booked',units_booked],['Excess Units', excess_units_received]
                ,['Short Units', short_units_received],['Units Recieved', units_received]]
                #val_df = val_df.set_index
                bar_df = pd.DataFrame(bar_data, columns=['Label', 'Units'])  
                #bar_df = bar_df.set_index('Label')       
                # bar_data = {
                #     'Units Booked':units_booked,
                #     'Excess Units Received': excess_units_received,
                #     'Short Units Received': short_units_received,
                #     'Units Recieved':units_received
                # }

                pie_data = [['Matching SKUs',matching_sku],['Mismatching SKUs', mismatching_sku]]
                pie_df = pd.DataFrame(pie_data, columns=['Label', 'Units']) 
                #pie_df = pie_df.set_index('Label')    
                with st.expander('Visualize Reconciliation Output'):
                    bar,pie = st.columns([1.2,1]) 
                    with bar:
                        plot_waterfall_chart(units_booked, excess_units_received, short_units_received, units_received)
                        #plot_bar_chart(bar_df,'Label','Units')
                        #st.bar_chart(bar_df)
                    with pie:
                        #st.bar_chart(pie_df)
                        plot_pie_chart(matching_sku, mismatching_sku)
            emp, but, empty = st.columns([2.05,1.2,1.5]) 
            with but:
                with open('temp/shipment_reco.xlsx', 'rb') as my_file:
                    click = st.download_button(label = 'Download in Excel', data = my_file, file_name = 'shipment_reco.xlsx', 
                    mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                #print(click) 
            #st.write(workbook) 
        else:
            if state.submit == True:
                if state.response != {}:
                    response = state.response
                    units_booked, excess_units_received, short_units_received, units_received, matching_sku, mismatching_sku = response
                    #print("Units :" + str(units_booked))
                    bar_data = [['Units Booked',units_booked],['Excess Units', excess_units_received]
                    ,['Short Units', short_units_received],['Units Recieved', units_received]]
                    #val_df = val_df.set_index
                    bar_df = pd.DataFrame(bar_data, columns=['Label', 'Units'])  
                    #bar_df = bar_df.set_index('Label')       
                    # bar_data = {
                    #     'Units Booked':units_booked,
                    #     'Excess Units Received': excess_units_received,
                    #     'Short Units Received': short_units_received,
                    #     'Units Recieved':units_received
                    # }

                    pie_data = [['Matching SKUs',matching_sku],['Mismatching SKUs', mismatching_sku]]
                    pie_df = pd.DataFrame(pie_data, columns=['Label', 'Units']) 
                    #pie_df = pie_df.set_index('Label')    
                    with st.expander('Visualize Reconciliation Output'):
                        bar,pie = st.columns([1.2,1]) 
                        with bar:
                            plot_waterfall_chart(units_booked, excess_units_received, short_units_received, units_received)
                            #plot_bar_chart(bar_df,'Label','Units')
                            #st.bar_chart(bar_df)
                        with pie:
                            #st.bar_chart(pie_df)
                            plot_pie_chart(matching_sku, mismatching_sku)
                    emp, but, empty = st.columns([2.05,1.2,1.5]) 
                    with but:
                        with open('temp/shipment_reco.xlsx', 'rb') as my_file:
                            click = st.download_button(label = 'Download in Excel', data = my_file, file_name = 'shipment_reco.xlsx', 
                            mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        # except:
        #     st.error("Run failed, kindly check if the inputs are valid")
                 
                
    
    def plot_bar_chart(data,X,Y):
        chart = (
            alt.Chart(data).configure_title(fontSize=20)
            .mark_bar()
            .encode(
                x=alt.X(X, type="nominal", title="", axis = alt.Axis(labelAngle=0,labelOverlap=False,labelAlign ='center',labelFontSize=10.5)),
                y=alt.Y(Y, type="quantitative", title=""),
                color = alt.Color(X, legend=None),
                # color=alt.condition(
                # alt.datum[Y] > 0,
                # alt.value("#74c476"),  # The positive color
                # alt.value("#d6616b")  # The negative color
                # ),
                tooltip = [alt.Tooltip(Y, title="",format='.1f')]
                #color=alt.Color("variable", type="nominal", title=""),
                #order=alt.Order("variable", sort="descending"),
            )
        ).interactive()
        
        st.altair_chart(chart, use_container_width=True)

    #@st.cache(suppress_st_warning=True, allow_output_mutation=True)
    def plot_waterfall_chart(units_booked, excess_units_received, short_units_received, units_received):
        fig = go.Figure(go.Waterfall(
            name = "20", orientation = "v",
            measure = ["relative", "relative", "relative", "total"],
            x = ["Units Booked", "Excess Units", "Short Units", "Units Recieved"],
            textposition = "outside",
            text = [str(units_booked), str(excess_units_received), str(short_units_received), str(units_received)],
            y = [units_booked, excess_units_received, short_units_received, units_received],
            decreasing = {"marker":{"color":"#002878"}},
            increasing = {"marker":{"color":"#00B0F0"}},
            totals = {"marker":{"color":"#002878"}},
            connector = {"line":{"color":"rgb(63, 63, 63)"}},
        ))

        fig.update_layout(
                title = 'Reconciliation Movement', title_x=0.5,
                showlegend = False, 
                yaxis=dict(showgrid=False),
                dragmode = "pan",
                selectdirection= "v"
        )
        fig.update_layout({
            'plot_bgcolor': 'rgba(0, 0, 0, 0)',
            'paper_bgcolor': 'rgba(0, 0, 0, 0)',
            })
        fig.update_traces(cliponaxis=False)

        config = dict({'displayModeBar': False,
                        'responsive': False,
                        'staticPlot': True
                        })
        st.plotly_chart(fig, use_container_width=True, config = config)
    
    #@st.cache(suppress_st_warning=True, allow_output_mutation=True)
    def plot_pie_chart(matching_sku, mismatching_sku):
        labels = ['Matching SKUs','Mismatching SKUs']
        values = [matching_sku, mismatching_sku]
        
        mark_colours = ['#00B0F0','#002878']

        fig = go.Figure(data=[go.Pie(labels=labels, values=values, hole=.3, marker_colors=mark_colours)])
        

        fig.update_layout(
                title = 'Reconciliation Split',title_x=0.48,title_y = 0.91,
                hovermode= False,
                legend_yanchor="bottom",legend_y = -0.15,legend_xanchor="center",legend_x = 0.5,
                legend_orientation= "h"
                #showlegend = False
        )
        config = dict({'displayModeBar': False,
                        'responsive': False
                        })

        st.plotly_chart(fig, use_container_width=True, config = config)

    def delete_temp():
        os.remove("temp/shipment_reco.xlsx")

    def file_upload_form():
        colour = "#89CFF0"
        with st.form(key = 'ticker',clear_on_submit=True):
            text, upload = st.columns([2.5,3]) 
            with text:
                st.write("###")
                st.write("###")
                st.write(f'<h5>{"&nbsp; Upload Shipment Instruction:"}</h5>', unsafe_allow_html=True)
            with upload:
                shipment_instructions = st.file_uploader("",key = 'ship_ins', accept_multiple_files=True)

            text, upload = st.columns([2.5,3])
            with text:
                st.write("###")
                st.write("###")
                st.write(f'<h5>{"&nbsp; Upload Warehouse Reports:"}<h5>', unsafe_allow_html=True)
            with upload:
                warehouse_reports = st.file_uploader("",key = 'ware_rep', accept_multiple_files=True)

            text, upload = st.columns([2.5,3])
            with text:
                st.write("###")
                st.write("###")
                st.write(f'<h5> {"&nbsp; Upload Inventory Ledger:"}<h5>', unsafe_allow_html=True)
            with upload:
                inventory_ledger = st.file_uploader("",key = 'inv_led')
            
            a,button,b = st.columns([2,1.2,1.5]) 
            with button:
                st.write('###')
                submit = st.form_submit_button(label = "Start Reconciliation")
                #submit = st.button(label="Start Reconciliation")
        return shipment_instructions, warehouse_reports, inventory_ledger, submit
        

        

    landing_page()