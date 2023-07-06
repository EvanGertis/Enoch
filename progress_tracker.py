import openpyxl

def prompt_user():
    top_priority = input("What is your top priority for the day? ")
    second_priority = input("What is your second priority for the day? ")
    third_priority = input("What is your third priority for the day? ")
    advice = input("Who can you call for advice on achieving your goals? ")
    actions = input("What specific actions do you need to take to move closer to your objectives? ")
    devotional = input("Did you take 10-15 minutes for morning devotional or meditation? ")
    workout = input("Did you complete a 30-45 minute workout or core exercise routine? ")
    budget_review = input("Did you complete your weekly budget review and stick to your budget for the week (assuming an average spending amount of $176.1 per category)? ")
    overspending = {}
    for category in ["House Taxes", "House Insurance", "Electricity", "House Gas", "Water/Garbage", "House Repairs", "Other Housing Costs", "Groceries", "Vehicle Gas", "Grandkids", "Vehicle Taxes and Insurance", "Vehicle Replacement", "Oil Changes", "AAA", "Tires", "Vehicle Repairs/Upkeep", "Health Insurance Premiums", "Healthcare Costs", "Cell Phones", "Mortgage", "Car Payments", "Other Debt Payments", "Giving", "Emergency Fund", "Brokerage Account", "Vacations", "Entertainment", "Restaurants", "Clothing", "Family", "Christmas", "Gifts", "Hobbies", "House Upgrades", "Cable", "Internet", "Subscriptions", "Lawn", "Beauty Products", "Gym Memberships", "Pets", "Life Insurance", "Disability Income Insurance", "Tax-Free Savings Bucket", "Medical Expenses/Hospital", "Lawyers", "Other"]:
        overspending[category] = input(f"Did you overspend in {category}? If so, how much?")
    unexpected_expenses = {}
    for category in ["House Taxes", "House Insurance", "Electricity", "House Gas", "Water/Garbage", "House Repairs", "Other Housing Costs", "Groceries", "Vehicle Gas", "Grandkids", "Vehicle Taxes and Insurance", "Vehicle Replacement", "Oil Changes", "AAA", "Tires", "Vehicle Repairs/Upkeep", "Health Insurance Premiums", "Healthcare Costs", "Cell Phones", "Mortgage", "Car Payments", "Other Debt Payments", "Giving", "Emergency Fund", "Brokerage Account", "Vacations", "Entertainment", "Restaurants", "Clothing", "Family", "Christmas", "Gifts", "Hobbies", "House Upgrades", "Cable", "Internet", "Subscriptions", "Lawn", "Beauty Products", "Gym Memberships", "Pets", "Life Insurance", "Disability Income Insurance", "Tax-Free Savings Bucket", "Medical Expenses/Hospital", "Lawyers", "Other"]:
        unexpected_expenses[category] = input(f"Were there any unexpected expenses that came up in {category}? If so, how much?")
    remaining_monthly_budget = {}
    for category in ["House Taxes", "House Insurance", "Electricity", "House Gas", "Water/Garbage", "House Repairs", "Other Housing Costs", "Groceries", "Vehicle Gas", "Grandkids", "Vehicle Taxes and Insurance", "Vehicle Replacement", "Oil Changes", "AAA", "Tires", "Vehicle Repairs/Upkeep", "Health Insurance Premiums", "Healthcare Costs", "Cell Phones", "Mortgage", "Car Payments", "Other Debt Payments", "Giving", "Emergency Fund", "Brokerage Account", "Vacations", "Entertainment", "Restaurants", "Clothing", "Family", "Christmas", "Gifts", "Hobbies", "House Upgrades", "Cable", "Internet", "Subscriptions", "Lawn", "Beauty Products", "Gym Memberships", "Pets", "Life Insurance", "Disability Income Insurance", "Tax-Free Savings Bucket", "Medical Expenses/Hospital", "Lawyers", "Other"]:
        remaining_monthly_budget[category] = input(f"How much money do you have left for the rest of the month in {category}?")

    return top_priority, second_priority, third_priority, advice, actions, devotional, workout, budget_review, overspending, unexpected_expenses, remaining_monthly_budget

def create_excel_worksheet():
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = "Daily Progress Tracker"
    sheet["A1"] = "Daily Progress Tracker"
    sheet["A2"] = "Date:"
    sheet["B2"] = "Goals for the Day"
    sheet["C2"] = "Advice from Mentor"
    sheet["D2"] = "Actions to Take"
    sheet["E2"] = "Morning Devotional/Meditation"
    sheet["F2"] = "Workout/Core Exercise Routine"
    sheet["G2"] = "Budget Review"
    sheet["H2"] = "Overspending"
    sheet["AI2"] = "Unexpected Expenses"
    sheet["AS2"] = "Remaining Monthly Budget"

    # Set column widths
    sheet.column_dimensions["A"].width = 25
    sheet.column_dimensions["B"].width = 25
    sheet.column_dimensions["C"].width = 25
    sheet.column_dimensions["D"].width = 25
    sheet.column_dimensions["E"].width = 25
    sheet.column_dimensions["F"].width = 25
    sheet.column_dimensions["G"].width = 25

    # Write data to worksheet
    top_priority, second_priority, third_priority, advice, actions, devotional, workout, budget_review, overspending, unexpected_expenses, remaining_monthly_budget = prompt_user()

    sheet["B4"] = top_priority
    sheet["B5"] = second_priority
    sheet["B6"] = third_priority
    sheet["C4"] = advice
    sheet["D4"] = actions
    sheet["E4"] = devotional
    sheet["F4"] = workout
    sheet["G4"] = budget_review

    # Write overspending data to worksheet
    for i, category in enumerate(overspending.keys()):
        sheet[f"H{i+4}"] = overspending[category]

    # Write unexpected expenses data to worksheet
    for i, category in enumerate(unexpected_expenses.keys()):
        sheet[f"AI{i+4}"] = unexpected_expenses[category]

    # Write remaining monthly budget data to worksheet
    for i, category in enumerate(remaining_monthly_budget.keys()):
        sheet[f"AS{i+4}"] = remaining_monthly_budget[category]

    wb.save("daily_progress_tracker.xlsx")

if __name__ == "__main__":
    create_excel_worksheet()