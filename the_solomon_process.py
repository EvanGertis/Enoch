import openpyxl

def prompt_user():
    top_priority = float(input("What is your top priority for the day? "))
    second_priority = float(input("What is your second priority for the day? "))
    third_priority = float(input("What is your third priority for the day? "))
    advice = input("Who can you call for advice on achieving your goals? ")
    actions = input("What specific actions do you need to take to move closer to your objectives? ")
    devotional = input("Did you take 10-15 minutes for morning devotional or meditation? ")
    workout = input("Did you complete a 30-45 minute workout or core exercise routine? ")
    budget_review = input("Did you complete your weekly budget review and stick to your budget for the week (assuming an average spending amount of $176.1 per category)? ")
    overspending = {}
    for category in ["House Taxes", "House Insurance", "Electricity", "House Gas", "Water/Garbage", "House Repairs", "Other Housing Costs", "Groceries", "Vehicle Gas", "Grandkids", "Vehicle Taxes and Insurance", "Vehicle Replacement", "Oil Changes", "AAA", "Tires", "Vehicle Repairs/Upkeep", "Health Insurance Premiums", "Healthcare Costs", "Cell Phones", "Mortgage", "Car Payments", "Other Debt Payments", "Giving", "Emergency Fund", "Brokerage Account", "Vacations", "Entertainment", "Restaurants", "Clothing", "Family", "Christmas", "Gifts", "Hobbies", "House Upgrades", "Cable", "Internet", "Subscriptions", "Lawn", "Beauty Products", "Gym Memberships", "Pets", "Life Insurance", "Disability Income Insurance", "Tax-Free Savings Bucket", "Medical Expenses/Hospital", "Lawyers", "Other"]:
        overspending[category] = float(input(f"Did you overspend in {category}? If so, how much? "))
    unexpected_expenses = {}
    for category in ["House Taxes", "House Insurance", "Electricity", "House Gas", "Water/Garbage", "House Repairs", "Other Housing Costs", "Groceries", "Vehicle Gas", "Grandkids", "Vehicle Taxes and Insurance", "Vehicle Replacement", "Oil Changes", "AAA", "Tires", "Vehicle Repairs/Upkeep", "Health Insurance Premiums", "Healthcare Costs", "Cell Phones", "Mortgage", "Car Payments", "Other Debt Payments", "Giving", "Emergency Fund", "Brokerage Account", "Vacations", "Entertainment", "Restaurants", "Clothing", "Family", "Christmas", "Gifts", "Hobbies", "House Upgrades", "Cable", "Internet", "Subscriptions", "Lawn", "Beauty Products", "Gym Memberships", "Pets", "Life Insurance", "Disability Income Insurance", "Tax-Free Savings Bucket", "Medical Expenses/Hospital", "Lawyers", "Other"]:
        unexpected_expenses[category] = float(input(f"Were there any unexpected expenses that came up in {category}? If so, how much? "))
    remaining_monthly_budget = {}
    for category in ["House Taxes", "House Insurance", "Electricity", "House Gas", "Water/Garbage", "House Repairs", "Other Housing Costs", "Groceries", "Vehicle Gas", "Grandkids", "Vehicle Taxes and Insurance", "Vehicle Replacement", "Oil Changes", "AAA", "Tires", "Vehicle Repairs/Upkeep", "Health Insurance Premiums", "Healthcare Costs", "Cell Phones", "Mortgage", "Car Payments", "Other Debt Payments", "Giving", "Emergency Fund", "Brokerage Account", "Vacations", "Entertainment", "Restaurants", "Clothing", "Family", "Christmas", "Gifts", "Hobbies", "House Upgrades", "Cable", "Internet", "Subscriptions", "Lawn", "Beauty Products", "Gym Memberships", "Pets", "Life Insurance", "Disability Income Insurance", "Tax-Free Savings Bucket", "Medical Expenses/Hospital", "Lawyers", "Other"]:
        remaining_monthly_budget[category] = float(input(f"How much money do you have left for the rest of the month in {category}? "))

    return top_priority, second_priority, third_priority, advice, actions, devotional, workout, budget_review, overspending, unexpected_expenses, remaining_monthly_budget

def create_excel_worksheet():
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = "Daily Progress Tracker"
    sheet["A1"] = "Daily Progress Tracker"
    sheet["A2"] = "Date:"
    sheet["B2"] = "7/5/2023"
    sheet["A4"] = "Top Priority"
    sheet["B4"] = "Second Priority"
    sheet["C4"] = "Third Priority"
    sheet["D4"] = "Advice"
    sheet["E4"] = "Actions"
    sheet["F4"] = "Devotional/Meditation"
    sheet["G4"] = "Workout/Core Exercise"
    sheet["H4"] = "Budget Review"
    sheet["I4"] = "Overspending"
    sheet["I5"] = "Category"
    sheet["J5"] = "Amount"
    sheet["AI4"] = "Unexpected Expenses"
    sheet["AI5"] = "Category"
    sheet["AJ5"] = "Amount"
    sheet["BA4"] = "Remaining Monthly Budget"
    sheet["BA5"] = "Category"
    sheet["BB5"] = "Amount"

    # prompt user for data and insert it into the worksheet
    top_priority, second_priority, third_priority, advice, actions, devotional, workout, budget_review, overspending, unexpected_expenses, remaining_monthly_budget = prompt_user()
    sheet["A5"] = top_priority
    sheet["B5"] = second_priority
    sheet["C5"] = third_priority
    sheet["D5"] = advice
    sheet["E5"] = actions
    sheet["F5"] = devotional
    sheet["G5"] = workout
    sheet["H5"] = budget_review
    for i, category in enumerate(overspending.keys()):
        sheet[f"I{i+6}"] = category
        sheet[f"J{i+6}"] = overspending[category]
    for i, category in enumerate(unexpected_expenses.keys()):
        sheet[f"AI{i+6}"] = category
        sheet[f"AJ{i+6}"] = unexpected_expenses[category]
    for i, category in enumerate(remaining_monthly_budget.keys()):
        sheet[f"BA{i+6}"] = category
        sheet[f"BB{i+6}"] = remaining_monthly_budget[category]

    # save the workbook
    wb.save("daily_progress_tracker.xlsx")

if __name__ == "__main__":
    create_excel_worksheet()