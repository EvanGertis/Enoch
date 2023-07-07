import openpyxl
from openpyxl.utils.exceptions import InvalidFileException
import webbrowser

CATEGORIES = {
    "Emergency Fund": 0.1,
    "Brokerage Account": 0.05,
    "Vacations": 0.05,
    "Entertainment": 0.08,
    "Restaurants": 0.05,
    "Clothing": 0.038,
    "Family": 0.05,
    "Christmas": 0.02,
    "Gifts": 0.02,
    "Hobbies": 0.1,
    "House Upgrades": 0.05,
    "Cable": 0.02,
    "Internet": 0.02,
    "Subscriptions": 0.02,
    "Lawn": 0.02,
    "Beauty Products": 0.02,
    "Gym Memberships": 0.02,
    "Mortgage": 0.18,
    "House Taxes": 0.05,
    "House Insurance": 0.02,
    "Electricity": 0.18,
    "House gas": 0.02,
    "Water / Garbage": 0.02,
    "House Repairs": 0.05,
    "Other housing costs": 0.05,
    "Groceries": 0.1,
    "Vehicle Gas": 0.02,
    "Grandkids": 0.02,
    "Vehicle Taxes and Insurance": 0.05,
    "Vehicle Replacement": 0.05,
    "Oil Changes": 0.02,
    "AAA": 0.02,
    "Tires": 0.02,
    "Vehicle Repairs / Upkeep": 0.05,
    "Health Insurance Premiums": 0.1,
    "Healthcare Costs": 0.05,
    "Cell Phones": 0.05,
    "Credit Card Payments": 0.05,
    "CarPayments": 0.05,
}

def open_urls():
    urls = [
        "https://docs.google.com/spreadsheets/d/1ZQtV99mXJm-MogabMc0fuR5B9IQ8arv8/edit#gid=1344099930",
        "https://logon.vanguard.com",
        "https://robinhood.com",
        "https://blockfi.com",
        "https://mint.intuit.com/",
        "https://next.waveapps.com/",
        "https://www.ncsecu.org",
        "https://www.bankofamerica.com",
        "https://www.paypal.com",
        "https://www.advisorclient.com",
        "https://client.schwab.com",
        "https://dashboard.stripe.com",
        "https://www.buymeacoffee.com/dashboard",
        "https://www.udemy.com/instructor/performance/overview/revenue/?date_filter=year&data_scope=all",
        "https://www.linkedin.com/"
    ]
    
    for url in urls:
        webbrowser.open(url)

def prompt_user():
    top_priority = input("What is your top priority for the day? ")
    second_priority = input("What is your second priority for the day? ")
    third_priority = input("What is your third priority for the day? ")
    advice = input("Who can you call for advice on achieving your goals? ")
    actions = input("What specific actions do you need to take to move closer to your objectives? ")
    devotional = input("Did you take 10-15 minutes for morning devotional or meditation? (yes/no) ")
    workout = input("Did you complete a 30-45 minute workout or core exercise routine? (yes/no) ")
    budget_review = input("Did you complete your weekly budget review and stick to your budget for the week (assuming an average spending amount of $176.1 per category)? (yes/no) ")
    overspending = {}
    for category in CATEGORIES:
        response = input(f"Did you overspend in {category}? If so, how much? (enter a number or '0' if not applicable) ")
        while not (response.isdigit() or response == "0"):
            response = input(f"Invalid input. Did you overspend in {category}? If so, how much? (enter a number or '0' if not applicable) ")
        overspending[category] = float(response)
    unexpected_expenses = {}
    for category in CATEGORIES:
        response = input(f"Were there any unexpected expenses that came up in {category}? If so, how much? (enter a number or '0' if not applicable) ")
        while not (response.isdigit() or response == "0"):
            response = input(f"Invalid input. Were there any unexpected expenses that came up in {category}? If so, how much? (enter a number or '0' if not applicable) ")
        unexpected_expenses[category] = float(response)
    
    # Calculate the remaining monthly budget for each category based on income and percentages
    remaining_monthly_budget = {}
    total_income = 1174  # Assuming $1174 per week
    for category, percentage in CATEGORIES.items():
        remaining_monthly_budget[category] = round(total_income * percentage * 4, 2)
        if category in overspending:
            remaining_monthly_budget[category] -= overspending[category]
        if category in unexpected_expenses:
            remaining_monthly_budget[category] -= unexpected_expenses[category]

    return (top_priority, second_priority, third_priority, advice, actions, devotional, workout, budget_review, overspending, unexpected_expenses, remaining_monthly_budget)

def write_to_excel(file_path, data):
    try:
        workbook = openpyxl.load_workbook(file_path)
    except InvalidFileException:
        workbook = openpyxl.create_workbook(file_path)

    sheet = workbook.active
    sheet["A1"] = "Top Priority"
    sheet["B1"] = "Second Priority"
    sheet["C1"] = "Third Priority"
    sheet["D1"] = "Advice"
    sheet["E1"] = "Actions"
    sheet["F1"] = "Devotional/Meditation"
    sheet["G1"] = "Workout/Core Exercise"
    sheet["H1"] = "Budget Review"
    sheet["I1"] = "Overspending"
    sheet["J1"] = "Unexpected Expenses"
    sheet["K1"] = "Remaining Monthly Budget"

    row_num = sheet.max_row + 1
    sheet[f"A{row_num}"] = data[0]
    sheet[f"B{row_num}"] = data[1]
    sheet[f"C{row_num}"] = data[2]
    sheet[f"D{row_num}"] = data[3]
    sheet[f"E{row_num}"] = data[4]
    sheet[f"F{row_num}"] = data[5]
    sheet[f"G{row_num}"] = data[6]
    sheet[f"H{row_num}"] = data[7]
    sheet[f"I{row_num}"] = "\n".join([f"{category}: {amount}" for category, amount in data[8].items()])
    sheet[f"J{row_num}"] = "\n".join([f"{category}: {amount}" for category, amount in data[9].items()])
    sheet[f"K{row_num}"] = "\n".join([f"{category}: {amount}" for category, amount in data[10].items()])

    workbook.save(file_path)

def open_budget_spreadsheet(file_path):
    try:
        webbrowser.open(file_path)
    except:
        print("Unable to open the budget spreadsheet. Please check the file path and try again.")

if __name__ == "__main__":
    open_urls()
    # data = prompt_user()
    # write_to_excel("daily_report.xlsx", data)
    # open_budget_spreadsheet("budget.xlsx")