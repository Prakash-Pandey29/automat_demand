import streamlit as st
import pandas as pd
from datetime import datetime

# Streamlit App Title
st.title("Excel Data Analysis App")

# Date Input
date_input = st.text_input("Enter a date (e.g., August 2023):")

# File Upload
st.subheader("Upload Excel Files")
excel_file1 = st.file_uploader(
    "Upload the All Demands_Data Dump Excel file", type=["xlsx", "xls"]
)
excel_file2 = st.file_uploader(
    "Upload the Mapping Excel file", type=["xlsx", "xls"]
)

# Check if the Run Code button is clicked
if st.button("Run Code"):
    # Validate Date Input
    try:
        date_obj = datetime.strptime(date_input, "%B %Y")
    except ValueError:
        st.error(
            "Invalid date format. Please enter a date in the format 'Month Year' (e.g., August 2023)."
        )
        st.stop()

    # Check if Excel files are uploaded
    if excel_file1 is not None and excel_file2 is not None:

        #import your file
        # demand data loading
        demand_data = pd.read_excel(excel_file1, engine='openpyxl')

        # mapping data loading
        map_data = pd.read_excel(excel_file2, engine='openpyxl')

        # saving the different mappings for BU in a dictionary
        map_dictionary = {}
        for name in map_data['BU'].unique():
            map_dictionary[name] = map_data[map_data.BU==name]['BUOps'].iloc[0]

        # dates are in string format so converting this into datetime format
        demand_data['Status Date'] = pd.to_datetime(demand_data['Status Date'], errors = 'coerce')
        demand_data['Demand From']= pd.to_datetime(demand_data['Demand From'], errors = 'coerce')
        demand_data['Ops Planned Date']= pd.to_datetime(demand_data['Ops Planned Date'], errors = 'coerce')
        demand_data['Billability Start Date']= pd.to_datetime(demand_data['Billability Start Date'], errors = 'coerce')

        # creating month column
        demand_data['Status_month'] = demand_data['Status Date'].dt.strftime('%B %Y')
        demand_data['Demand_from_month'] = demand_data['Demand From'].dt.strftime('%B %Y')
        demand_data['Ops_date_month']= demand_data['Ops Planned Date'].dt.strftime('%B %Y')
        demand_data['B_start_month'] = demand_data['Billability Start Date'].dt.strftime('%B %Y')


        # creating a new column "BUOps" and mapping values from BU column
        demand_data['BUOps'] = None
        for val in map_dictionary:
            demand_data.loc[(demand_data.BU==val), 'BUOps'] = map_dictionary[val]

        #demand type = new
        demand_data = demand_data[demand_data['Demand Type']=='New']

        # Allocation Sub Type = Billable only
        demand_data = demand_data[demand_data['Allocation Sub Type']=='Billable']

        # removing internal demands that starts with RE means those has been extended
        demand_data = demand_data[~demand_data['Demand No.'].str.startswith('RE')]

        # removing cancelled or hold status demand
        demand_data = demand_data[~demand_data['Status'].isin(['Canceled', 'Hold', 'Open'])]

        # filter on ramp up reasons
        demand_data = demand_data[~demand_data['Ramp Up Reason'].isin(['P2P', "SOW Extension", 'Grace Extension (awaiting SOW Renewal)', 'Billable to Investment/NB'])]

        # filter on ramp down reason
        demand_data = demand_data[~demand_data['Ramp Down Reason'].isin(['InCorrect Demand'])]

        # demand_data1 = pd.read_excel('All Demands_Data Dump.xlsx', engine='openpyxl')
        demand_data[demand_data['Ops Planned Date']==demand_data['Ops Planned Date'].min()]['Ramp Down Reason'].value_counts(dropna=False)

        # input the desired month (e.g August 2023 )
        month = date_input

        #creating the reporting status column below:

        # Aug fullfiled : Status = Fulfiled & Status month = Aug 23 (regardless of BSD month)
        # D Univ : Status != Fulfiled
        # Prior Fulfilled, Aug BSD : Status = Fulfiled & BSD month = Aug 23
        # what about : Prior Fulfilled, BSD not in Aug, May be in sep, or october
        month_fulfilled =  f'{month[:3]} fulfilled'
        demand_data['Reporting_status'] = None
        demand_data.loc[(demand_data.Status=='Fulfilled') & (demand_data.Status_month==month), 'Reporting_status'] = month_fulfilled
        demand_data.loc[(demand_data.Status=='Fulfilled') & (demand_data.Status_month!=month) & (demand_data.B_start_month==month), 'Reporting_status'] = f'Prior Fulfilled, {month[:3]} BSD'
        demand_data.loc[(demand_data.Status!='Fulfilled'), 'Reporting_status'] = 'D Universe'

        # billing start month before input month (ex:aug), status date before aug, delete :
        filter_date = datetime.strptime(month, '%B %Y')
        demand_data['B_start_month'] = pd.to_datetime(demand_data['B_start_month'], format='%B %Y')
        demand_data['Status_month'] = pd.to_datetime(demand_data['Status_month'], format='%B %Y')
        # demand_data = demand_data[~((demand_data.B_start_month<filter_date) & (demand_data.Status_month<filter_date))]
        demand_data = demand_data[~((demand_data.B_start_month<filter_date) & (demand_data.Status_month<filter_date) & ((demand_data.Status=="Fulfilled")))]

        #changing type of the column into string
        demand_data['Status_month'] = demand_data['Status Date'].dt.strftime('%B %Y')
        demand_data['B_start_month'] = demand_data['B_start_month'].dt.strftime('%B %Y')


        # creating a new column Stream1 from stream and sub-stream
        demand_data['Stream1'] = None
        demand_data.loc[(demand_data['Sub Stream']=='Big Data Engineering'), 'Stream1'] = 'DE'
        demand_data.loc[(demand_data['Sub Stream']=='Data Warehouse'), 'Stream1'] = 'DE'
        demand_data.loc[(demand_data['Sub Stream']=='Quality Engineering'), 'Stream1'] = 'DE'
        demand_data.loc[(demand_data['Sub Stream']=='Business Intelligence'), 'Stream1'] = 'DE'
        demand_data.loc[(demand_data['Sub Stream']=='Data Science'), 'Stream1'] = 'DS'
        demand_data.loc[(demand_data['Sub Stream']=='Application Engineering'), 'Stream1'] = 'APP & Design'
        demand_data.loc[(demand_data['Sub Stream']=='MLE'), 'Stream1'] = 'MLE'
        demand_data.loc[(demand_data['Sub Stream']=='Devops'), 'Stream1'] = 'DevOps'
        demand_data.loc[(demand_data['Sub Stream']=='Data Science Python'), 'Stream1'] = 'DS'
        demand_data.loc[(demand_data['Sub Stream']=='MLOPS'), 'Stream1'] = 'MLE'
        demand_data.loc[(demand_data['Sub Stream']=='Data Science Insights'), 'Stream1'] = 'DS'

        demand_data.loc[(demand_data['Sub Stream'].isna()) & (demand_data.Stream=='Analytics Consulting'), 'Stream1'] = 'AC'
        demand_data.loc[(demand_data['Sub Stream'].isna()) & (demand_data.Stream=='Technology Consulting'), 'Stream1'] = 'TC'
        demand_data.loc[(demand_data['Sub Stream'].isna()) & (demand_data.Stream=='Design Consulting'), 'Stream1'] = 'APP & Design'
        demand_data.loc[(demand_data['Sub Stream'].isna()) & (demand_data.Stream=='Production Support'), 'Stream1'] = 'Prod Support'
        demand_data.loc[(demand_data['Sub Stream'].isna()) & (demand_data.Stream=='Data Science'), 'Stream1'] = 'DS'


        #FTE Calculation
        # storing all the months related column in a list
        allocation_months = [col for col in demand_data.columns if 'Allocation %' in col]

        #reversing the order of the months
        allocation_months = allocation_months[::-1]

        for col in allocation_months:
            demand_data[col] = demand_data[col].str.rstrip('%').astype(float)

        # creating a new column name FTE with some conditions related to months column
        demand_data['FTE'] = 0
        for col in allocation_months:
            demand_data.loc[(demand_data.FTE==0), 'FTE'] = demand_data[col]/100

        #Pivot Table
        input_ops_planned_date = 'January 2000'

        pivot_table_input_month_fullfiled_DU_univ = demand_data[(demand_data.Reporting_status==month_fulfilled) | (demand_data.Reporting_status=='D Universe')]

        #excluding Jan 2000'January 2000'
        pivot_table_input_month_fullfiled_DU_univ_without_Jan00 = demand_data[(demand_data.Reporting_status==month_fulfilled) | (demand_data.Reporting_status=='D Universe') & (demand_data['Ops Planned Date']!=input_ops_planned_date)]

        #first pivot table
        pivot_demand_universe = pivot_table_input_month_fullfiled_DU_univ.pivot_table(index=['BUOps','Customer Name'] ,values='FTE', columns='Stream1', aggfunc='sum')
        pivot_demand_universe = pivot_demand_universe.assign(DU_Total=pivot_demand_universe.sum(axis=1)).reset_index()

        #second pivot table
        pivot_demand_universe_with_ops_p_d = pivot_table_input_month_fullfiled_DU_univ_without_Jan00.pivot_table(index=['BUOps','Customer Name'] ,values='FTE', columns='Stream1', aggfunc='sum')
        pivot_demand_universe_with_ops_p_d = pivot_demand_universe_with_ops_p_d.assign(DU_Ops_Total=pivot_demand_universe_with_ops_p_d.sum(axis=1)).reset_index().reset_index()

        #droping index col
        pivot_demand_universe_with_ops_p_d.drop(['index'], axis=1, inplace=True)

        # this is multiindexing here (if anyone want multiindexing, then he can use below code)

        # pivot_demand_universe.columns = pd.MultiIndex.from_tuples([('Demand Universe', col) for col in pivot_demand_universe.columns])
        # pivot_demand_universe_with_ops_p_d.columns = pd.MultiIndex.from_tuples([('DEMAND UNIVERSE with OPS PLAN DATE', col) for col in pivot_demand_universe_with_ops_p_d.columns])

        #merge both pivot table
        merged_pivot = pd.merge(pivot_demand_universe, pivot_demand_universe_with_ops_p_d, on=['BUOps', 'Customer Name'], suffixes=('_DU', "_DU_with_OPD"))

        # prior fulfilled , input month bsd
        prior_fulfulled_input_month_bsd = demand_data[demand_data.Reporting_status==f'Prior Fulfilled, {month[:3]} BSD']
        prior_fulfulled_input_month_bsd = prior_fulfulled_input_month_bsd.groupby('Customer Name', as_index=False)['FTE'].sum()
        prior_fulfulled_input_month_bsd.rename({'FTE':f'{month[:3]} BSD'}, inplace=True, axis=1)

        result = pd.merge(merged_pivot, prior_fulfulled_input_month_bsd, on="Customer Name", how="left")



        #       # # buffer to use for excel writer
        # import io
        # buffer = io.BytesIO()
        # # download button 2 to download dataframe as xlsx
        # with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        #     # Write each dataframe to a different worksheet.
        #     result.to_excel(writer, sheet_name='Sheet1', index=False)

        #     download2 = st.download_button(
        #         label="Download data as Excel",
        #         data=buffer,
        #         file_name='result_button.xlsx',
        #         mime='application/vnd.ms-excel')
        st.download_button(
        label="Download data as CSV",
        data=result.to_csv(index=False).encode('utf-8'),
        file_name='result_2.csv',
        mime='text/csv',
)

        # For example, you can display some statistics or merge the dataframes

        # # Display the result
        # st.subheader("Data Analysis Result:")
        # st.write("Date Input:", date_obj.strftime("%B %Y"))
        # st.write("Summary of Excel File 1:")
        # st.write(df1.head())
        # st.write("Summary of Excel File 2:")
        # st.write(df2.head())

    else:
        st.warning("Please upload both Excel files.")

# Optionally, you can add more features and data analysis code based on your requirements.
