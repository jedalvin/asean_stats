import pandas as pd
import plotly.graph_objects as go
import matplotlib.pyplot as plt
import streamlit as st
import datetime
import altair as alt

st.set_page_config(layout='wide')

# Streamlit app
st.title('PSEi vs Other Asian Exchanges')

# File upload
uploaded_file = st.file_uploader("Upload the stock market index Excel file", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)

    # Function to extract 'Date' and 'Close' columns from a sheet, handling non-breaking spaces
    def extract_date_close(sheet_name):
        df = pd.read_excel(xls, sheet_name)
        df.columns = df.iloc[0]  # Set the first row as the column names
        df = df[1:]  # Skip the first row
        df.columns = df.columns.str.strip()  # Remove leading/trailing whitespace from column names
        df = df[['Date', 'Close']]  # Extract 'Date' and 'Close' columns
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')  # Convert 'Date' to datetime, handle errors gracefully
        df.dropna(subset=['Date'], inplace=True)  # Drop rows where 'Date' is NaN after conversion
        df.set_index('Date', inplace=True)  # Set 'Date' as index
        df['Close'] = pd.to_numeric(df['Close'], errors='coerce')  # Convert 'Close' to numeric, handle errors gracefully
        df.dropna(subset=['Close'], inplace=True)  # Drop rows where 'Close' is NaN after conversion
        return df

    # Plot using Plotly
    fig = go.Figure()

    # Define the plotting function for indices
    def plot_indices(combined_df, selected_indices=None):
        for column in combined_df.columns:
            if selected_indices and column not in selected_indices:
                continue
            color = '#02025f' if column == 'PSEI' else 'lightgray'
            fig.add_trace(go.Scatter(x=combined_df.index, y=combined_df[column], mode='lines', name=column, line=dict(color=color)))

    # Function to plot bar charts with annotations
    def plot_bar_chart(data, title):
        fig, ax = plt.subplots(figsize=(10, 6))
        for index, row in data.iterrows():
            color = '#02025f' if row['Index'] != 'PSEI' else '#F6C324'
            ax.barh(row['Index'], row['YTD Change'], color=color)
        ax.set_title(title)
        ax.set_xlabel('YTD Change (%)')
        ax.set_ylabel('Index')
        ax.invert_yaxis()  # Invert y-axis to show rank 1 at the top
        plt.tight_layout()
        return fig

    # List all sheet names
    sheet_names = xls.sheet_names

    # List of relevant sheet names (excluding 'Legend', 'stock_index', and 'Summary')
    relevant_sheets = sheet_names[1:]

    # Extract and combine data from all relevant sheets
    combined_df = pd.DataFrame()

    for sheet in relevant_sheets:
        try:
            sheet_df = extract_date_close(sheet)
            sheet_df.rename(columns={'Close': sheet}, inplace=True)  # Rename 'Close' column to sheet name
            combined_df = pd.concat([combined_df, sheet_df], axis=1)  # Combine data
        except Exception as e:
            st.error(f"Error processing sheet '{sheet}': {e}")

    # Define the fixed indices to be shown for ASEAN
    asean_indices = ['VNI', 'KLCI', 'PSEI', 'STI', 'JAKIDX', 'SET']

    # Select box to choose between ASEAN indices and all indices
    option = st.selectbox('Select Indices', ['ASEAN Indices', 'All Indices'])

    # Display historical closing prices plot
    st.subheader('Plot of Historical Closing Prices')

    # Plot based on the selected option
    fig = go.Figure()
    if option == 'ASEAN Indices':
        plot_indices(combined_df, asean_indices)
    else:
        plot_indices(combined_df)

    # Update layout for plot
    fig.update_layout(
        xaxis_title='Year',
        yaxis_title='Closing Price',
        legend_title='Index',
        template='plotly_white'
    )

    # Display the plot using Streamlit in full width
    st.plotly_chart(fig, use_container_width=True)

    # User input date
    input_date = st.date_input("Enter the date:", datetime.date(2024, 6, 25))

    # List to store results for the three tables
    results_list = []
    results_list_prev_month = []
    results_list_two_year = []

    # Process each relevant sheet
    for sheet in relevant_sheets:
        try:
            sheet_df = extract_date_close(sheet)

            # Calculate Table 1: Year-to-Date Change based on User Input Date
            fiscal_year_end = pd.to_datetime(datetime.datetime(input_date.year - 1, 12, 31).date())

            last_date_prev_year = sheet_df[sheet_df.index.year == fiscal_year_end.year].index.max()
            year_end_price = sheet_df.loc[last_date_prev_year]['Close'] if last_date_prev_year else None

            input_date_price = sheet_df.loc[pd.Timestamp(input_date)]['Close'] if pd.Timestamp(input_date) in sheet_df.index else sheet_df[sheet_df.index <= pd.Timestamp(input_date)].iloc[-1]['Close']

            if year_end_price is not None:
                ytd_change = round(((input_date_price - year_end_price) / year_end_price) * 100, 2)
                results_list.append({
                    'Index': sheet,
                    'Year-End Closing Price': year_end_price,
                    'Input Date Closing Price': input_date_price,
                    'YTD Change': ytd_change
                })

            # Calculate Table 2: Year-to-Date Change based on Previous Month-End

            # Last date of the previous month
            prev_month_end = (pd.Timestamp(input_date).replace(day=1) - pd.Timedelta(days=1)).replace(day=1) + pd.offsets.MonthEnd(1)

            try:
                prev_month_end_price = sheet_df.loc[prev_month_end]['Close']
            except KeyError:
                prev_month_end_price = None

            if prev_month_end_price is not None:
                ytd_change_prev_month = round(((prev_month_end_price - year_end_price) / year_end_price) * 100, 2)
                results_list_prev_month.append({
                    'Index': sheet,
                    'Year-End Closing Price': year_end_price,
                    'Previous Month-End Closing Price': prev_month_end_price,
                    'YTD Change': ytd_change_prev_month
                })

            # Calculate Table 3: Year-to-Date Change based on Last Two Year-End Closing Prices  
            two_preceding_years = fiscal_year_end.year - 1
            last_date_two_preceding_years = sheet_df[sheet_df.index.year == two_preceding_years].index.max()
            two_preceding_years_end_price = sheet_df.loc[last_date_two_preceding_years]['Close'] if last_date_two_preceding_years else None

            if two_preceding_years_end_price is not None:
                ytd_change_two_year = round(((year_end_price - two_preceding_years_end_price) / two_preceding_years_end_price) * 100, 2)
                results_list_two_year.append({
                    'Index': sheet,
                    'Previous Year-End Closing Price': two_preceding_years_end_price,
                    'Year-End Closing Price': year_end_price,
                    'YTD Change': ytd_change_two_year
                })
        except Exception as e:
            st.error(f"Error processing sheet '{sheet}': {e}")

    # Convert results to DataFrames
    results_df = pd.DataFrame(results_list)
    results_prev_month_df = pd.DataFrame(results_list_prev_month)
    results_two_year_df = pd.DataFrame(results_list_two_year)

    # Sort results by YTD Change
    results_df = results_df.sort_values(by='YTD Change', ascending=False)
    results_prev_month_df = results_prev_month_df.sort_values(by='YTD Change', ascending=False)
    results_two_year_df = results_two_year_df.sort_values(by='YTD Change', ascending=False)

    # Rank the indices
    results_df['Rank'] = results_df['YTD Change'].rank(ascending=False, method='min').astype(int)
    results_prev_month_df['Rank'] = results_prev_month_df['YTD Change'].rank(ascending=False, method='min').astype(int)
    results_two_year_df['Rank'] = results_two_year_df['YTD Change'].rank(ascending=False, method='min').astype(int)

    # Sort results by YTD Change
    results_df.set_index('Rank', inplace=True)
    results_prev_month_df.set_index('Rank', inplace=True)
    results_two_year_df.set_index('Rank', inplace=True)

    # Display Table 1, Table 2, and Table 3 side by side
    col1, col2, col3 = st.columns(3)

    # Filtered tables based on option selected
    if option == 'ASEAN Indices':
        results_df = results_df[results_df['Index'].isin(asean_indices)]
        results_prev_month_df = results_prev_month_df[results_prev_month_df['Index'].isin(asean_indices)]
        results_two_year_df = results_two_year_df[results_two_year_df['Index'].isin(asean_indices)]

    with col1:
        st.markdown(f"<h3 style='font-size:14px;'>Figure 1: YTD Change based on User Input Date ({input_date.strftime('%Y-%m-%d')})</h3>", unsafe_allow_html=True)
        # Example usage: Plotting Table 1: Year-to-Date Change based on User Input Date
        fig1 = plot_bar_chart(results_df, 'YTD Change based on User Input Date')

        # Display the plot using st.pyplot()
        st.pyplot(fig1)
        st.dataframe(results_df[['Index', 'Year-End Closing Price', 'Input Date Closing Price', 'YTD Change']])

    with col2:
        st.markdown("<h3 style='font-size:14px;'>Figure 2: Year-to-Date Change based on Previous Month-End Closing Price</h3>", unsafe_allow_html=True)
        # Example usage: Plotting Table 2: Year-to-Date Change based on Previous Month-End Closing Price
        fig2 = plot_bar_chart(results_prev_month_df, 'YTD Change based on Previous Month-End Closing Price')

        # Display the plot using st.pyplot()
        st.pyplot(fig2)
        st.dataframe(results_prev_month_df[['Index', 'Year-End Closing Price', 'Previous Month-End Closing Price', 'YTD Change']])

    with col3:
        st.markdown(f"<h3 style='font-size:14px;'>Figure 3: Change based on Last Two Year-End Closing Prices ({last_date_prev_year.year} vs. {two_preceding_years})</h3>", unsafe_allow_html=True)
        # Example usage: Plotting Table 3: Year-to-Date Change based on Last Two Year-End Closing Prices
        fig3 = plot_bar_chart(results_two_year_df, 'YTD Change based on Last Two Year-End Closing Prices')

        # Display the plot using st.pyplot()
        st.pyplot(fig3)
        st.dataframe(results_two_year_df[['Index', 'Previous Year-End Closing Price', 'Year-End Closing Price', 'YTD Change']])
else:
    st.info("Please upload an Excel file to proceed.")
