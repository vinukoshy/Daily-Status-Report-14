import pandas as pd
import openpyxl
import plotly.graph_objects as go
import plotly.express as px
import pandas as pd
from plotly.subplots import make_subplots
import numpy as np
from datetime import date
import io
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from PIL import Image
from reportlab.lib.utils import ImageReader
from PIL import Image
import streamlit as st
import os
import plotly.io as pio
from io import BytesIO


def image_combiner_sidebyside(jpeg_images):
    # check if number of images is odd, and if so, append a white image
    if len(jpeg_images) % 2 != 0:
        white_image = Image.new('RGB', (jpeg_images[0].width, jpeg_images[0].height), (255, 255, 255))
        jpeg_images.append(white_image)
        # combine JPEG images horizontally
    max_images_per_row = 2
    n_rows = -(-len(jpeg_images) // max_images_per_row)  # ceil division
    row_height = jpeg_images[0].height
    combined_image = Image.new('RGB', (max_images_per_row * jpeg_images[0].width, n_rows * row_height))
    x_offset = 0
    y_offset = 0
    for i, jpeg_image in enumerate(jpeg_images):
        combined_image.paste(jpeg_image, (x_offset, y_offset))
        x_offset += jpeg_image.width
        if (i + 1) % max_images_per_row == 0:
            x_offset = 0
            y_offset += row_height
    return combined_image

def image_combiner_below(jpeg_images):
    # calculate the total height of all the images combined
    total_height = sum([img.height for img in jpeg_images])

    # adjust the height of the images in combined_image
    combined_image = Image.new("RGB", (int(sum([img.width * 1.2 for img in jpeg_images])), int(jpeg_images[0].height * len(jpeg_images) * 1.2)), color='white')
    x_offset = (combined_image.width - jpeg_images[0].width)//2 # center align the images
    y_offset = 0
    for img in jpeg_images:
        img_width = int(img.width * 1.2)
        img_height = int(img.height * 1.2)
        combined_image.paste(img.resize((img_width, img_height)), (x_offset, y_offset))
        y_offset += img_height
    return combined_image

def sslc_manpower(dfs,sheets):
    sslcc_sheet_names = [sheet for sheet in sheets if sheet.startswith("SSLCC")]
    sslcc_details = [name.split("_")[1] for name in sslcc_sheet_names]
    dict_sslc = {}
    html_images =[]
    today = date.today().strftime("%B %d, %Y")
    for index, sslc_sheet in enumerate(sslcc_sheet_names):
        sslcc_subheader = sslcc_details[index]
        df_sslc = dfs[sslc_sheet]
        float_cols = df_sslc.select_dtypes(include='float').columns
        df_sslc[float_cols] = df_sslc[float_cols].fillna(0)
        df_sslc[float_cols] = df_sslc[float_cols].astype(int)
        jpeg_images =[]
        # get unique companies
        companies = df_sslc['Company'].unique()

        # loop through companies
        for company in companies:
            skip_flag = ""
            # subset data for current company
            company_data = df_sslc[df_sslc['Company'] == company]
            # calculate the total count of all subcategories for the current company
            total_count = company_data.iloc[:, 2:].sum().sum()
            # get unique categories for current company
            categories = company_data['Category'].unique().tolist()
            
        # check number of categories
            if len(categories) == 1:
                # if only one category, plot category text next to plot title
                fig = go.Figure()
                category = categories[0]
                category_data = company_data[company_data['Category'] == category].iloc[0, 2:]
                bar_data = [go.Bar(x=[category_data.index[i]], y=[category_data.values[i]], text=[category_data.values[i]], name=category_data.index[i], textposition='auto') for i in range(len(category_data)) if category_data[i] > 0]
                fig.add_traces(bar_data)
                fig.update_layout(title=f'<b>{company} - {category} (Total Count: {category_data.sum()})</b>')
                fig.update_yaxes(title_text="Count")
                fig.update_layout(legend={"title": "Designation"})
                if category_data.sum() == 0:
                    fig.update_layout(
                    xaxis=dict(visible=False),
                    yaxis=dict(visible=False),
                    showlegend = False,
                    plot_bgcolor='white',
                    paper_bgcolor='white')
                    skip_flag = "X"
            else:
                # if multiple categories, create subplots for each category
                fig = make_subplots(rows=len(categories), cols=1, subplot_titles=[f"{category} (Total Count: {company_data[company_data['Category'] == category].iloc[0, 2:].sum()})" for category in categories], vertical_spacing=0.65)
                for i, category in enumerate(categories):
                    category_data = company_data[company_data['Category'] == category].iloc[0, 2:]
                    bar_data = [go.Bar(x=[category_data.index[i]], y=[category_data.values[i]], text=[category_data.values[i]], name=category_data.index[i], textposition='auto') for i in range(len(category_data)) if category_data[i] > 0]
                    fig.add_traces(bar_data, rows=[i+1]*len(bar_data), cols=[1]*len(bar_data))
                    fig.update_yaxes(title_text="Count", row=i+1, col=1)
                    fig.update_layout(title=f'<b>{company} (Total Count: {company_data.iloc[:, 2:].sum().sum()})</b>')
                    fig.update_layout(legend={"title": "Designation"})
            fig.update_traces(marker_color=fig.layout.colorway)
            

            # convert Plotly figure object to JPEG
            fig_bytes = fig.to_image(format='jpg', width=800, height=600, scale=2)
            jpeg_image = Image.open(io.BytesIO(fig_bytes))
            jpeg_images.append(jpeg_image)
            if len(categories) == 1:
                    # set chart title and axis labels
                fig.update_layout(title=f'<b>SSLC {sslcc_subheader} {company} - {category} (Total Count: {category_data.sum()})</b> ({today})')
            else:
                fig.update_layout(title=f'<b>SSLC {sslcc_subheader} {company} (Total Count: {company_data.iloc[:, 2:].sum().sum()})</b> ({today})')
            if not skip_flag:
                html_images.append(fig)
        if len(jpeg_images) > 3:
            combined_image = image_combiner_sidebyside(jpeg_images)
        else:
            combined_image = image_combiner_below(jpeg_images)
        dict_sslc_report = {sslc_sheet : combined_image}
        dict_sslc.update(dict_sslc_report)
    return dict_sslc, html_images

def recruitment(dfs,sheets):
    rec_sheet_names = [sheet for sheet in sheets if sheet.capitalize().startswith("Recruitment")]
    rec_details = [name.split("_")[1] for name in rec_sheet_names ]
    dict_rec = {}
    jpeg_images =[]
    html_images =[]
    today = date.today().strftime("%B %d, %Y")
    for index, rec_sheet in enumerate(rec_sheet_names):
        rec_subheader = rec_details[index]
        df_rec = dfs[rec_sheet]
        float_cols = df_rec.select_dtypes(include='float').columns
        df_rec[float_cols] = df_rec[float_cols].fillna(0)
        df_rec[float_cols] = df_rec[float_cols].astype(int)
        df_rec.insert(0, 'Company', rec_subheader)

        # get unique companies
        companies = df_rec['Company'].unique()


        # loop through companies and create individual plots
        for company in companies:
            # filter dataframe for current company
            df_company = df_rec[df_rec["Company"] == company]
                # melt dataframe to long format
            df_melt = df_company.melt(id_vars=["Company", "Designation"], var_name="Status", value_name="Count")
            
            # create grouped bar chart with text labels
            fig = px.bar(df_melt, x="Designation", y="Count", color="Status", barmode="group",
                        text=[f"{row['Count']}" for _, row in df_melt.iterrows()],
                        category_orders={"Designation": df_rec["Designation"].values})
            
            # set chart title and axis labels
            fig.update_layout(title=f"<b>Recruitment Status {company}</b> ({today})",
                            xaxis_title="Designation",
                            yaxis_title="Count")
            fig.update_traces(marker_color=fig.layout.colorway)
            

            # convert Plotly figure object to JPEG
            fig_bytes = fig.to_image(format='jpg', width=800, height=600, scale=2)
            jpeg_image = Image.open(io.BytesIO(fig_bytes))
            jpeg_images.append(jpeg_image)
            html_images.append(fig)
    if len(jpeg_images) > 3:
        combined_image = image_combiner_sidebyside(jpeg_images)
    else:
        combined_image = image_combiner_below(jpeg_images)
    dict_rec_report = {rec_sheet : combined_image}
    dict_rec.update(dict_rec_report)
    return dict_rec, html_images
            
def crew_status(dfs,sheets):
    crew_sheet_names = [sheet for sheet in sheets if sheet.capitalize().startswith("Crew")]
    crew_details = [name.split("_")[1] for name in crew_sheet_names ]
    dict_crew = {}
    jpeg_images =[]
    html_images =[]
    today = date.today().strftime("%B %d, %Y")
    for index, crew_sheet in enumerate(crew_sheet_names):
        crew_subheader = crew_details[index]
        df_crew = dfs[crew_sheet]
        float_cols = df_crew.select_dtypes(include='float').columns
        df_crew[float_cols] = df_crew[float_cols].fillna(0)
        df_crew[float_cols] = df_crew[float_cols].astype(int)
        unnamed_index = [index for index,elem in enumerate(df_crew.columns) if elem.startswith('Unnamed')][0]
        df_crew_mobil = df_crew.iloc[:,unnamed_index+1:]
        df_crew_manpower = df_crew.iloc[:, :unnamed_index]
        df_crew_mobil = df_crew_mobil.dropna()

                # Create subplots with the required types
        fig = make_subplots(rows=2, cols=1, specs=[[{'type': 'pie'}], [{'type': 'bar'}]],vertical_spacing=0.3,
                            subplot_titles=['', f'{df_crew_mobil[df_crew_mobil.columns[0]].iloc[0]} Manpower Status'])

        # add pie chart to the first subplot
        fig.add_trace(
            go.Pie(
                labels=['Plan Count', 'Actual Count'],
                values=[df_crew_mobil['Plan'][0], df_crew_mobil['Actual'][0]],
                textinfo='label+value',
                hoverinfo='label+value',
                showlegend=True, legendgroup='pie',legendgrouptitle_text='<b>Mobilisation Status</b>'
            ),
            row=1, col=1
        )

        # add grouped bar chart to the second subplot
        for cols in df_crew_manpower.columns[1:]:
            fig.add_trace(go.Bar(name=cols, x=df_crew_manpower['Designation'], y=df_crew_manpower[cols], 
                                text=df_crew_manpower[cols], textposition='auto',showlegend=True, legendgroup='bar',legendgrouptitle_text='<b>Manpower Status</b>'),row=2, col=1)

        white_spaces = " " * 35
        text = 'Mobilisation Status'
        fig.update_xaxes(title_text='Designation', row=2, col=1)
        fig.update_yaxes(title_text='Count', row=2, col=1)
        # Update the layout of the figure
        fig.update_layout(title=f'<b>Crew Formation Status {crew_subheader}</b> ({today})<br><br>{white_spaces} {df_crew_mobil[df_crew_mobil.columns[0]].iloc[0]} {text}', margin=dict(t=200))
        # convert Plotly figure object to JPEG 
        fig_bytes = fig.to_image(format='jpg', width=800, height=600, scale=2)
        jpeg_image = Image.open(io.BytesIO(fig_bytes))
        jpeg_images.append(jpeg_image)
        if index == 0:
            white_spaces = " " * 88
        else:
            white_spaces = " " * 92
        # Update the layout of the figure
        fig.update_layout(title=f'<b>Crew Formation Status {crew_subheader}</b> ({today})<br><br>{white_spaces} {df_crew_mobil[df_crew_mobil.columns[0]].iloc[0]} {text}', margin=dict(t=200))
        html_images.append(fig)

    if len(jpeg_images) > 3:
        combined_image = image_combiner_sidebyside(jpeg_images)
    else:
        combined_image = image_combiner_below(jpeg_images)
    dict_crew_report = {crew_sheet : combined_image}
    dict_crew.update(dict_crew_report)
    return dict_crew, html_images

def visa_status(dfs,sheets):
    visa_status_sheet_names = [sheet for sheet in sheets if sheet.title().startswith("Visa Status")]
    visa_status_details = [name.split("_")[1] for name in visa_status_sheet_names ]
    dict_visa_status = {}
    jpeg_images =[]
    html_images =[]
    today = date.today().strftime("%B %d, %Y")
    for index, visa_status_sheet in enumerate(visa_status_sheet_names):
        visa_status_subheader = visa_status_details[index]
        df_visa_status = dfs[visa_status_sheet]
        float_cols = df_visa_status.select_dtypes(include='float').columns
        df_visa_status[float_cols] = df_visa_status[float_cols].fillna(0)
        df_visa_status[float_cols] = df_visa_status[float_cols].astype(int)
        df_visa_status.insert(0, 'Company', visa_status_subheader)
        # get list of unique companies
        companies = df_visa_status["Company"].unique()
        # loop through covisa_statusnies and create individual plots
        for company in companies:
            # filter dataframe for current company
            df_company = df_visa_status[df_visa_status["Company"] == company]
            
            # melt dataframe to long format
            df_melt = df_company.melt(id_vars=["Company", "Trade"], var_name="Status", value_name="Count")
            
            # create grouped bar chart with text labels
            fig = px.bar(df_melt, x="Trade", y="Count", color="Status", barmode="group",
                        text=[f"{row['Count']}" for _, row in df_melt.iterrows()],
                        category_orders={"Trade": df_visa_status["Trade"].values})
            
            # set chart title and axis labels
            today = date.today().strftime("%B %d, %Y")
            fig.update_layout(title=f"<b>Visa Status {company}</b> ({today})",
                            xaxis_title="Trade",
                            yaxis_title="Count")
            
            # convert Plotly figure object to JPEG
            fig_bytes = fig.to_image(format='jpg', width=800, height=600, scale=2)
            jpeg_image = Image.open(io.BytesIO(fig_bytes))
            jpeg_images.append(jpeg_image)
            html_images.append(fig)
    if len(jpeg_images) > 3:
        combined_image = image_combiner_sidebyside(jpeg_images)
    else:
        combined_image = image_combiner_below(jpeg_images)
    dict_visa_status_report = {visa_status_sheet : combined_image}
    dict_visa_status.update(dict_visa_status_report)
    return dict_visa_status, html_images

def visa_transfer(dfs,sheets):
    visatransfer_sheet_names = [sheet for sheet in sheets if sheet.title().startswith("Visa Transfer")]
    jpeg_images =[]
    html_images =[]
    dict_visatransfer ={}
    today = date.today().strftime("%B %d, %Y")
    for visatransfer_sheet in visatransfer_sheet_names:
        df_visatransfer = dfs[visatransfer_sheet]
        float_cols = df_visatransfer.select_dtypes(include='float').columns
        df_visatransfer[float_cols] = df_visatransfer[float_cols].fillna(0)
        df_visatransfer[float_cols] = df_visatransfer[float_cols].astype(int)

    # loop over unique companies
    for company in df_visatransfer['Company'].unique():
        # filter dataframe for current company
        df_visatransfer_filtered = df_visatransfer[df_visatransfer['Company'] == company]
        # Keep columns where value in first row is non-zero
        nonzero_cols = df_visatransfer_filtered.loc[:, (df_visatransfer_filtered != 0).any(axis=0)].columns
        # Select only the desired columns
        df_visatransfer_filtered = df_visatransfer_filtered[nonzero_cols]
        # loop over unique categories for current company
        for category in df_visatransfer_filtered['Category'].unique():
            # filter dataframe for current company and category
            df_visatransfer_filtered_cat = df_visatransfer_filtered[df_visatransfer_filtered['Category'] == category]
            
            # calculate waterfall values
            result = np.concatenate(([df_visatransfer_filtered_cat.iloc[0].values[2]], df_visatransfer_filtered_cat.iloc[0].values[3:] * -1))
            
            x = df_visatransfer_filtered_cat.columns[2:]
            y = result
            
            fig = go.Figure(go.Waterfall(
                name = "Fall in Employee Count",
                orientation = "v",
                measure = ["absolute"] + ["relative"] * (len(y) - 1),
                x = x,
                y = y,
                connector = {"line":{"color":"rgb(63, 63, 63)"}},
                textposition = "inside",
                text = [str(abs(value)) for value in y],
            ))
            today = date.today().strftime("%B %d, %Y")
            fig.update_layout(
                title=f"<b>Visa Transfer plan {company} {category}</b> ({today})<br><br><b>Total Employees:</b> {y[0]} ",
                xaxis_title="Date",
                yaxis_title="Count",
                showlegend=False,)
            # convert Plotly figure object to JPEG
            fig_bytes = fig.to_image(format='jpg', width=800, height=600, scale=2)
            jpeg_image = Image.open(io.BytesIO(fig_bytes))
            jpeg_images.append(jpeg_image)
            html_images.append(fig)
    if len(jpeg_images) > 3:
        combined_image = image_combiner_sidebyside(jpeg_images)
    else:
        combined_image = image_combiner_below(jpeg_images)
    dict_visatransfer_report = {visatransfer_sheet : combined_image}
    dict_visatransfer.update(dict_visatransfer_report)
    return dict_visatransfer, html_images

def create_pdf_report(dict_final):
    # define the header style
    header_style = {
    'font': 'Helvetica-Bold',
    'fontsize': 15,
    'underline': True,
    'align': 'CENTER'}
    # create a new PDF with a title
    # Create a BytesIO object to store the PDF content as bytes
    buffer = BytesIO()
    today = date.today().strftime("%B %d, %Y")
    file_name = f"Daily Status Report({today}).pdf"
    c = canvas.Canvas(buffer, pagesize=letter)
    c.setTitle("Daily report({today})")
    c.setPageCompression(3)
    # draw the first page with image1 and text1
    
    selected_data = [data.split("_")[0].strip() for data in dict_final.keys()]
    my_dict = {elem: None for elem in selected_data}
    # Convert the dictionary back to a list
    selected_data = list(my_dict.keys())
    for selected_data in selected_data:
        result = {k: v for k, v in dict_final.items() if k.startswith(selected_data)}
        c.setPageCompression(3)
        c.setFont(header_style["font"], header_style["fontsize"])
        c.drawCentredString(letter[0] / 2, 750, f"{selected_data} Detailed Report as on {today}")
        for combined_image in result.values():
            image_reader = ImageReader(combined_image)
            c.drawImage(image_reader, 0, 0, width=letter[0], height=(letter[1] - header_style["fontsize"] - inch))
            c.showPage()
    c.save()
    # Save the PDF content to the BytesIO object
    pdf_bytes = buffer.getvalue()
    
    return file_name, pdf_bytes

def create_html(html_images_main):
    html_file_name = f'Daily Status Report({date.today().strftime("%B %d, %Y")}).html'
    html_str = ''
    for fig in html_images_main:
        html_str += fig.to_html(full_html=False, include_plotlyjs='cdn')

    # Create a BytesIO object to store the HTML content as bytes
    html_bytes = html_str.encode()
    return html_file_name, html_bytes

@st.cache
def load_images():
    return Image.open("dash.jpeg"), os.path.join(os.path.expanduser("~"), "Downloads")

@st.cache
def get_app_data(file):
    error_flag = ""
    if not file:
        st.error('🚩️Error: Please select a valid Excel file')
        error_flag = "X"
    if not error_flag:
        # Load the Excel workbook        
        pio.renderers.default = "iframe_connected"
        workbook = openpyxl.load_workbook(file)
        # Get all the sheet names in the workbook
        sheet_names = workbook.sheetnames
        df_all = pd.read_excel(file, sheet_name=sheet_names,header=2)
                
        # Final objects to store html and image data 
        html_images_main = []
        dict_final = {}
                
        # SSLC
        dict_sslc_main,sslc_html_images = sslc_manpower(df_all,sheet_names)
        dict_final.update(dict_sslc_main)
        html_images_main.extend(sslc_html_images)
                
        # Recruitment
        dict_rec_main,rec_html_images = recruitment(df_all,sheet_names)
        dict_final.update(dict_rec_main)
        html_images_main.extend(rec_html_images)
            
        # Crew
        dict_crew_main,crew_html_images = crew_status(df_all,sheet_names)
        dict_final.update(dict_crew_main)
        html_images_main.extend(crew_html_images)

        # Visa Status
        dict_visa_status_main,visa_status_html_images = visa_status(df_all,sheet_names)
        dict_final.update(dict_visa_status_main)
        html_images_main.extend(visa_status_html_images)

        # Visa Transfer
        dict_visa_transfer_main,visatransfer_html_images = visa_transfer(df_all,sheet_names)
        dict_final.update(dict_visa_transfer_main)
        html_images_main.extend(visatransfer_html_images)    
        pdf_file_name,pdf_bytes = create_pdf_report(dict_final)
        html_file_name, html_bytes = create_html(html_images_main)    
        return pdf_file_name,pdf_bytes, html_file_name, html_bytes

def process_run():
    st.write('---')
    st.subheader("Select the Daily status Report Excel")
    # Browse and select XML file
    # Add a file uploader to select the Excel file
    file = st.file_uploader('Select an Excel file', type=['xlsx', 'xls'])
    # Validate the file path selected by the user
    st.write('---')
    if file:
        with st.spinner('Please Wait ⌛...Generating Report📝'):
            pdf_file_name,pdf_bytes, html_file_name, html_bytes = get_app_data(file)
            st.success("Daily Status Reports Generated Successfully")
            st.subheader("Select the Below Buttons to Download the Reports 👇:")
            if st.download_button(label="Download HTML Report",
            data=BytesIO(html_bytes).getvalue(),
            file_name=html_file_name,
            mime="text/html",
            ):
                st.success(f"**{html_file_name}** saved sucessfully")
            if st.download_button(
            label="Download PDF Report",
            data=pdf_bytes,
            file_name=pdf_file_name,
            mime="application/pdf",
            ):
                st.success(f"**{pdf_file_name}** saved sucessfully")
    return None

st.set_page_config(page_title='Daily Status Report',page_icon ="📝",layout="wide")
st.title("Daily 🗓️ Status Report 📊 Generation App")
image,down_path = load_images()
st.image(image, width=200)
process_run()

