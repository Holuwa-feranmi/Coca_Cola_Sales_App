import pandas as pd 
import streamlit as st 
import altair as alt 

@st.cache_data
def load_data():
    df = pd.read_excel("Cocacola.xlsx", skiprows=4)
    df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')
    df.rename(columns={'invoice_date': 'date'}, inplace=True)

    if df.shape[1] > 0:
        delete_first_columns = df.columns[0]
        df = df.drop(columns=[delete_first_columns], axis=1)

    if df.shape[-1] > 5:
        columns_to_drop = df.columns[-5:].tolist()
        df = df.drop(columns=columns_to_drop, axis=1)

    
        st.title("COCACOLA SALES APP")
        st.write(df)
        return df

try:
    df = load_data()

    st.sidebar.header("Data Preview")
    st.sidebar.dataframe(df.head())

    if hasattr(df, 'columns') and "total_sales" in df.columns and "operating_profit" in df.columns:
        total_sales = float(df["total_sales"].sum())
        total_profit = float(df["operating_profit"].sum())
        overall_margin = (total_profit / total_sales * 100) if total_sales != 0 else 0

        st.header("1: Overall Sales Performance")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Sales", f"${total_sales:,.2f}")
        col2.metric("Total Operating Profit", f"${total_profit:,.2f}")
        col3.metric("Overall Margin", f"{overall_margin:.2f}%")
    else:
        st.warning("Cannot calculate metrics - check DataFrame and columns")
        st.subheader("Detailed Results")
        st.header("1: Overall Sales Performance")
        st.write(f"**Total Sales:** ${total_sales:,.2f}")
        st.write(f"**Total Profit:** ${total_profit:,.2f}")
        st.write(f"**Profit Margin:** {margin:.1f}%")


    # --- Section : Which beverage_brand is the most popular---
    st.header("2. Top 5 Beverage Brands By Total Sales")
    if "beverage_brand" in df.columns and "total_sales" in df.columns and "operating_profit" in df.columns and "units_sold" in df.columns:
        brand_sales = df.groupby('beverage_brand')['total_sales'].sum().sort_values(ascending=False).head(5)
        st.write(brand_sales)
        brand_sales_df = brand_sales.reset_index()  # Key step: Convert to DataFrame
    
        # Create Altair chart
        chart_1= alt.Chart(brand_sales_df).mark_bar().encode(
            x=alt.X('total_sales:Q', title='Total Sales ($)'),
            y=alt.Y('beverage_brand:N', sort='-x', title='Beverage Brand'),
            color=alt.Color('total_sales:Q', scale=alt.Scale(scheme='reds')),
            tooltip=[
                alt.Tooltip('beverage_brand:N', title='Brand'),
                alt.Tooltip('total_sales:Q', format='$,.0f', title='Sales')]
        ).properties(
                title='Top 5 Beverage Brands by Total Sales',
                width=600,
                height=300
        )
    
        st.altair_chart(chart_1, use_container_width=True)

        st.header("3. Top 5 Beverage Brands By Profit")
        brand_profit = df.groupby('beverage_brand')['operating_profit'].sum().sort_values(ascending=False).head(5)
        st.write(brand_profit)

        brand_profit_df = brand_profit.reset_index()

        # Create Altair chart
        chart_2= alt.Chart(brand_profit_df).mark_bar().encode(
            x=alt.X('operating_profit:Q', title='Operating Profit ($)'),
            y=alt.Y('beverage_brand:N', sort='-x', title='Beverage Brand'),
            color=alt.Color('operating_profit:Q', scale=alt.Scale(scheme='viridis')),
            tooltip=[
                alt.Tooltip('beverage_brand:N', title='Brand'),
                alt.Tooltip('operating_profit:Q', format='$,.0f', title='Profit')]
        ).properties(
                title='Brand Profitability Analysis',
                width=600,
                height=400
        )

        st.altair_chart(chart_2, use_container_width=True)

        st.header("4. Top 5 Beverage Brands By Units Sold")
        Unit_sold = df.groupby("beverage_brand")["units_sold"].sum().sort_values(ascending=False).head(5)
        st.write(Unit_sold)

        Unit_sold_df = Unit_sold.reset_index()
        chart_3= alt.Chart(Unit_sold_df).mark_bar().encode(
            x=alt.X("units_sold:Q", title="Unit Sold ($)"),
            y=alt.Y("beverage_brand:N", sort="-x", title="Beverage Brand"),
            color=alt.Color("units_sold:Q", scale=alt.Scale(scheme="blues")),
            tooltip=[
                alt.Tooltip("beverage_brand:N", title="Brand"),
                alt.Tooltip("units_sold:Q", format="$,.0f", title="Unit Sold")
            ]
        ).properties(
            title="Top Brand By Unit Sold",
            width=600,
            height=400
        )
        st.altair_chart(chart_3, use_container_width=True)

        # Section 3: Which region/state/city is performing best
        st.header("5. Sales And Profit By Region")
        if "region" in df.columns and "total_sales" in df.columns and "operating_profit" in df.columns:
            region_performance = df.groupby("region").agg(
                total_sales =("total_sales", "sum"),
                Total_profit =("operating_profit", "sum")
        ).sort_values(by='total_sales', ascending=False)
        st.write(region_performance)
        region_performance_df = region_performance.reset_index()
        chart_4= alt.Chart(region_performance_df).mark_bar().encode(
            x=alt.X("total_sales:Q", title="Total Sales ($)"),
            y=alt.Y("region:N", sort="-x", title="Region"),
            color=alt.Color("total_sales:Q", scale=alt.Scale(scheme="plasma")),
            tooltip=[
            alt.Tooltip("region:N", title="Region"),
            alt.Tooltip("total_sales:Q", format="$,.0f", title="Sales"),
            alt.Tooltip("total_profit:Q", format="$,.0f", title="Profit")
            ]
        ).properties(
            title="Sales And Profit By Region",
            width=600,
            height=400
        )
        st.altair_chart(chart_4, use_container_width=True)

        st.header("6. Sales And Profit By State")
        if "state" in df.columns and "total_sales" in df.columns and "operating_profit" in df.columns:
            state_performance = df.groupby("state").agg(
                total_sales = ("total_sales", "sum"),
                total_profit = ("operating_profit", "sum")
            ).sort_values(by="total_sales", ascending=False)
            st.write(state_performance.head(10))

            state_performance_df = state_performance.reset_index()

            chart_5 = alt.Chart(state_performance_df ).mark_bar().encode(
                x=alt.X("total_sales:Q", title="Total Sales ($)"),
                y=alt.Y("state:N", sort="-x", title="State"),
                color=alt.Color("total_sales:Q", scale=alt.Scale(scheme='viridis')),
                tooltip=[
                    alt.Tooltip("state:N", title="State"),
                    alt.Tooltip("total_sales", format="$,.0f", title="Sales"),
                    alt.Tooltip("total_profit", format="$,.0f", title="Profit")
                ]
            ).properties(
                title="Sales And Profit By State",
                width=600,
                height=400
            )
            st.altair_chart(chart_5, use_container_width=True)

        st.header("6. Sales And Profit By State")
        if "city" in df.columns and "total_sales" in df.columns and "operating_profit" in df.columns:
            city_performance = df.groupby("city").agg(
                total_sales = ("total_sales", "sum"),
                total_profit = ("operating_profit", "sum")
            ).sort_values(by="total_sales", ascending=False)
            st.write(state_performance.head(10))

            city_performance_df = city_performance.reset_index()

            chart_6 = alt.Chart(city_performance_df ).mark_bar().encode(
                x=alt.X("total_sales:Q", title="Total Sales ($)"),
                y=alt.Y("city:N", sort="-x", title="City"),
                color=alt.Color("total_sales:Q", scale=alt.Scale(scheme='cividis')),
                tooltip=[
                    alt.Tooltip("city:N", title="City"),
                    alt.Tooltip("total_sales", format="$,.0f", title="Sales"),
                    alt.Tooltip("total_profit", format="$,.0f", title="Profit")
                ]
            ).properties(
                title="Sales And Profit By City",
                width=600,
                height=400
            )
            st.altair_chart(chart_6, use_container_width=True)

            # Section 4:What is the average price per unit and units sold
            st.header("7. Average Price Per Unit And Unit Sold")
            if "price_per_unit" in df.columns and "units_sold" in df.columns:
                Avg_price_per_unit =df["price_per_unit"].mean()
                Avg_units_sold =df["units_sold"].mean()

                col1, col2 = st.columns(2)
    
                with col1:
                    st.metric(
                        label="💰 Average Price per Unit",
                        value=f"${Avg_price_per_unit:.2f}"
                    )
                with col2:
                    st.metric(
                        label="📦 Average Units Sold",
                        value=f"{Avg_units_sold:.0f} units"
                    )

                # Section 5: How do sales and profit trend over time (monthly and yearly)
                st.header("8.  Monthly Trend By Sales And Profit")
                if "date" in df.columns and "total_sales" in df.columns and "operating_profit" in df.columns:
                    Monthly_Trend = df.set_index("date").resample("M").agg(
                        total_sales=("total_sales", "sum"),
                        total_profit=("operating_profit", "sum")
                    )
                    st.write(Monthly_Trend)

                st.header("9. Overall Sales And Profit Of The Year")
                if "date" in df.columns and "total_sales" in df.columns and "operating_profit" in df.columns:
                    Yearly_Trend = df.set_index("date").resample('Y').agg(
                        total_sales =("total_sales", "sum"),
                        Total_profit =("operating_profit", "sum")        
                    )
                    st.write(Yearly_Trend)



except FileNotFoundError:
    st.error("The file was not found. Please check the file path.")
except pd.errors.EmptyDataError:
    st.error("The file is empty. Please check the file contents.")
except Exception as e:
    st.error(f"An error occurred: {e}")