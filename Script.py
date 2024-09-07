import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Initial data and assumptions
start_users = 10000  # Starting with 10,000 users
increment_users = 5000  # Increment of users for each step

# Updated assumptions for impressions
session_length = 10  # Session length in minutes
sessions_per_month = 1  # Sessions per month (average shopping trips)
ads_per_minute = 1  # Ads shown per minute
ad_fill_rate = 0.80  # Ad fill rate
ads_per_impression = 1  # Number of ads shown per impression

# Affiliate marketing assumptions
affiliate_ctr = 0.05  # 5% click-through rate
affiliate_conversion_rate = 0.03  # 3% conversion rate
affiliate_aov = 50  # Average Order Value $
affiliate_commission_rate = 0.0175  # Weighted Commission Rate 1.75%

# Function to calculate weighted eCPM based on ad mix
def calculate_weighted_ecpm():
    ad_mix = {
        "Banner Ads": {"percentage": 0.33, "ecpm": 2.00},
        "Video Ads": {"percentage": 0.33, "ecpm": 15.00},
        "User-Initiated Ads": {"percentage": 0.33, "ecpm": 20.00},
    }
    weighted_ecpm = sum(ad["percentage"] * ad["ecpm"] for ad in ad_mix.values())
    return weighted_ecpm

weighted_ecpm = calculate_weighted_ecpm()

# Update impressions per user per month calculation
impressions_per_user_per_month = session_length * sessions_per_month * ads_per_minute

# Updated Administrative Costs (monthly)
admin_costs = 1000  # Updated monthly administrative costs
equipment_amortized = 100  # Amortized equipment costs per month

# Total Initial Monthly Expenses (excluding salaries)
initial_expenses = admin_costs + equipment_amortized  # Initial monthly expenses excluding CEO salary

initial_ceo_salary = 1000  # Initial CEO monthly salary
max_ceo_salary = 10000  # Maximum CEO monthly salary
churn_rate = 0.08  # Assuming an average churn rate of 8%
premium_conversion_rate = 0.03  # Premium conversion rate

# Assumptions for Salaries (annual)
ceo_annual_salary = 100000
senior_dev_annual_salary = 100000
junior_dev_annual_salary = 100000
designer_annual_salary = 100000
marketing_annual_salary = 100000

# Monthly Salaries
ceo_monthly_salary = ceo_annual_salary / 12
senior_dev_monthly_salary = senior_dev_annual_salary / 12
junior_dev_monthly_salary = junior_dev_annual_salary / 12
designer_monthly_salary = designer_annual_salary / 12
marketing_monthly_salary = marketing_annual_salary / 12

# Function to calculate subscription revenue per user
def calculate_subscription_revenue():
    subscription_distribution = {
        "Minimum Payment": 0.85,
        "Mid-range Payment": 0.10,
        "Maximum Payment": 0.05
    }
    subscription_prices = {
        "Minimum Payment": 0.99,
        "Mid-range Payment": 0.99,
        "Maximum Payment": 0.99
    }
    revenue = 0
    for tier, percentage in subscription_distribution.items():
        price = subscription_prices[tier]
        revenue += percentage * price
    return revenue

# Assumptions for COGS
server_costs_per_user = 0.016  # Server costs per user per month
bandwidth_costs_per_user = 0.009  # Bandwidth costs per user per month
other_infrastructure_costs_per_user = 0.028  # Other infrastructure costs per user per month

# Total COGS per user
total_cogs_per_user = server_costs_per_user + bandwidth_costs_per_user + other_infrastructure_costs_per_user

# Lists to store projected data
user_counts = []
ad_revenues = []
subscription_revenues = []
affiliate_revenues = []  # New list for Affiliate Revenue
total_revenues = []
cogs = []
operating_expenses = []
app_store_fees = []
net_incomes = []
arpu = []
arpu_growth = []  # New list for ARPU growth
clv = []
rpu = []
cpu = []
ppu = []
impressions_per_user_per_month_list = []
impressions_per_user_per_year_list = []
ad_ecpm_list = []
revenue_per_non_premium_user_per_year_list = []
revenue_per_premium_user_per_year_list = []  # Combined for First $1M and Above $1M
weighted_avg_revenue_per_user_per_year_list = []  # Combined for First $1M and Above $1M
session_length_list = []
sessions_per_month_list = []
ads_per_minute_list = []
ad_fill_rate_list = []
ads_per_impression_list = []
employee_counts = []
ceo_salaries = []

# Calculation by subscriber count
current_users = start_users
current_expenses = initial_expenses + initial_ceo_salary
current_employee_count = 1  # Starting with the CEO only
current_ceo_salary = initial_ceo_salary
hired_positions = ["CEO"]
positions_salaries = [("Senior Developer", senior_dev_monthly_salary), ("Junior Developer", junior_dev_monthly_salary),
                      ("Designer", designer_monthly_salary), ("Marketing", marketing_monthly_salary)]

for i in range(100):  # Limit to 100 rows for manageability
    user_counts.append(current_users)
    employee_counts.append(current_employee_count)
    ceo_salaries.append(current_ceo_salary)
    
    # Calculate Non-Premium User Count
    non_premium_users = round(current_users * (1 - premium_conversion_rate))
    
    # Calculate Premium User Count
    premium_users = round(current_users * premium_conversion_rate)
    
    # Calculate Ad Revenue
    ad_revenue = non_premium_users * impressions_per_user_per_month * (weighted_ecpm / 1000) * ad_fill_rate * ads_per_impression
    ad_revenues.append(round(ad_revenue, 2))
    
    # Calculate Subscription Revenue
    subscription_revenue = premium_users * calculate_subscription_revenue()
    subscription_revenues.append(round(subscription_revenue, 2))
    
    # Calculate Affiliate Revenue
    affiliate_clicks = current_users * affiliate_ctr
    affiliate_conversions = affiliate_clicks * affiliate_conversion_rate
    affiliate_revenue = affiliate_conversions * affiliate_aov * affiliate_commission_rate
    affiliate_revenues.append(round(affiliate_revenue, 2))
    
    # Ensure all revenue components are correctly summed in total_revenue
    total_revenue = ad_revenue + subscription_revenue + affiliate_revenue
    total_revenues.append(round(total_revenue, 2))
    
    # Calculate App Store Fees
    cumulative_revenue = sum(total_revenues)
    if cumulative_revenue <= 1000000:
        app_store_fee = subscription_revenue * 0.15 + affiliate_revenue * 0.15
    else:
        app_store_fee = subscription_revenue * 0.30 + affiliate_revenue * 0.30
    app_store_fees.append(round(app_store_fee, 2))
    
    # Calculate Net Income
    net_income = total_revenue - (current_users * total_cogs_per_user) - current_expenses - app_store_fee
    net_incomes.append(round(net_income, 2))
    
    # Calculate ARPU (Average Revenue Per User)
    if current_users > 0:
        arpu_value = total_revenue / current_users
    else:
        arpu_value = 0
    arpu.append(round(arpu_value, 2))
    
    # Calculate ARPU Growth (%)
    if i == 0:
        arpu_growth_value = 0  # No growth for the first data point
    else:
        arpu_growth_value = ((arpu[-1] - arpu[-2]) / arpu[-2]) * 100 if arpu[-2] != 0 else 0
    arpu_growth.append(round(arpu_growth_value, 2))
    
    # Calculate CLV (Customer Lifetime Value)
    clv_value = arpu_value / churn_rate if churn_rate != 0 else 0
    clv.append(round(clv_value, 2))
    
    # Calculate RPU (Revenue Per User)
    rpu_value = total_revenue / current_users if current_users > 0 else 0
    rpu.append(round(rpu_value, 2))
    
    # Calculate Total Costs
    total_costs = (current_users * total_cogs_per_user) + current_expenses + app_store_fee
    
    # Calculate CPU (Cost Per User)
    cpu_value = total_costs / current_users if current_users > 0 else 0
    cpu.append(round(cpu_value, 2))
    
    # Calculate PPU (Profit Per User)
    ppu_value = rpu_value - cpu_value
    ppu.append(round(ppu_value, 2))
    
    # Fixed values
    cogs.append(round(current_users * total_cogs_per_user, 2))
    operating_expenses.append(round(current_expenses, 2))
    impressions_per_user_per_month_list.append(impressions_per_user_per_month)
    impressions_per_user_per_year_list.append(impressions_per_user_per_month * 12)
    ad_ecpm_list.append(weighted_ecpm)
    revenue_per_non_premium_user_per_year_list.append(round(impressions_per_user_per_month * weighted_ecpm / 1000 * ad_fill_rate * ads_per_impression * 12, 2))

    # Session and ad parameters
    session_length_list.append(session_length)
    sessions_per_month_list.append(sessions_per_month)
    ads_per_minute_list.append(ads_per_minute)
    ad_fill_rate_list.append(ad_fill_rate)
    ads_per_impression_list.append(ads_per_impression)

    # Weighted average revenue calculations
    premium_user_revenue = calculate_subscription_revenue() * 12  # Annualize the subscription revenue
    revenue_per_premium_user_per_year_list.append(round(premium_user_revenue, 2))
    weighted_avg_revenue_per_user_per_year_list.append(round(0.97 * (impressions_per_user_per_month * weighted_ecpm / 1000 * ad_fill_rate * ads_per_impression * 12) + 0.03 * premium_user_revenue, 2))

    # Increment user count
    current_users += increment_users

    # Update current expenses and onboard employees as needed
    if i > 0:  # Skip the first row as it's initial setup
        if net_income > current_expenses + 10000 and positions_salaries:
            position, salary = positions_salaries.pop(0)
            current_expenses += salary
            current_employee_count += 1
            hired_positions.append(position)

    # Scale CEO Salary
    if net_income > current_ceo_salary + 10000 and current_ceo_salary < max_ceo_salary:
        current_ceo_salary = min(current_ceo_salary + 1000, max_ceo_salary)

# Create DataFrame with formulas
data = {
    "User Count": user_counts,
    "Ad Revenue ($)": ad_revenues,
    "Subscription Revenue ($)": subscription_revenues,
    "Affiliate Revenue ($)": affiliate_revenues,  # New column for Affiliate Revenue
    "Total Revenue ($)": total_revenues,
    "COGS ($)": cogs,
    "Operating Expenses ($)": operating_expenses,
    "App Store Fees ($)": app_store_fees,
    "Net Income ($)": net_incomes,
    "ARPU ($) (Average Revenue Per User)": arpu,
    "ARPU Growth (%)": arpu_growth,  # New column for ARPU growth
    "CLV ($) (Customer Lifetime Value)": clv,
    "RPU ($) (Revenue Per User)": rpu,  # New column for Revenue per User
    "CPU ($) (Cost Per User)": cpu,  # New column for Cost per User
    "PPU ($) (Profit Per User)": ppu,  # New column for Profit per User
    "Impressions per User per Month": impressions_per_user_per_month_list,
    "Impressions per User per Year": impressions_per_user_per_year_list,
    "Ad eCPM ($)": ad_ecpm_list,
    "Revenue per Non-Premium User per Year ($)": revenue_per_non_premium_user_per_year_list,
    "Revenue per Premium User per Year ($)": revenue_per_premium_user_per_year_list,
    "Weighted Average Revenue per User per Year ($)": weighted_avg_revenue_per_user_per_year_list,
    "Session Length (minutes)": session_length_list,
    "Sessions per Month": sessions_per_month_list,
    "Ads per Minute": ads_per_minute_list,
    "Ad Fill Rate (%)": ad_fill_rate_list,
    "Number of Ads Shown per Impression": ads_per_impression_list,
    "Total Employee Count": employee_counts,
    "CEO Salary ($)": ceo_salaries
}

# Ensure all lists are the same length
max_length = max(map(len, data.values()))
for key in data:
    if len(data[key]) < max_length:
        data[key] += [data[key][-1]] * (max_length - len(data[key]))

df = pd.DataFrame(data)

# Create Excel file with embedded formulas using openpyxl
file_path_with_formulas = r"YOURPATHHERE-CHANGE ME"
with pd.ExcelWriter(file_path_with_formulas, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name='Financial Projections')
    workbook = writer.book
    worksheet = writer.sheets['Financial Projections']

    for row in dataframe_to_rows(df, index=False, header=True):
        worksheet.append(row)

    for row in worksheet.iter_rows(min_row=2, max_row=len(df) + 1, min_col=1, max_col=len(df.columns)):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                cell.value = cell.value

# Save the workbook
workbook.save(file_path_with_formulas)

