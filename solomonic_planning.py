import openpyxl

def financial_questions():
    money_made_today = input("How much moneydid you make today? ")
    money_spent_today = input("How much money did you spend today? ")
    money_saved_today = input("How much money did you save today? ")
    money_invested_today = input("How much money did you invest today? ")
    return money_made_today, money_spent_today, money_saved_today, money_invested_today

def prompt_user():
    top_priority = input("What is your top priority for the day? ")
    financial_data = financial_questions()
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
    total_income = float(input("What is your total income for the month? "))
    remaining_monthly_budget = {}
    for category, percentage in [("Emergency Fund", 10), ("Brokerage Account", 5), ("Vacations", 5), ("Entertainment", 8), ("Restaurants", 5), ("Clothing", 3.8), ("Family", 5), ("Christmas", 2), ("Gifts", 2), ("Hobbies", 10), ("House Upgrades", 5), ("Cable", 2), ("Internet", 2), ("Subscriptions", 2), ("Lawn", 2), ("Beauty Products", 2), ("Gym Memberships", 2), ("Mortgage", 18), ("House Taxes", 5), ("House Insurance", 2), ("Electricity", 18), ("House gas", 2), ("Water / Garbage", 2), ("House Repairs", 5), ("Other housing costs", 5), ("Groceries", 10), ("Vehicle Gas", 2), ("Grandkids", 2), ("Vehicle Taxes and Insurance", 5), ("Vehicle Replacement", 5), ("Oil Changes", 2), ("AAA", 2), ("Tires", 2), ("Vehicle Repairs / Upkeep", 5), ("Health Insurance Premiums", 10), ("Healthcare Costs", 5), ("Cell Phones", 5), ("Credit Card Payments", 5), ("CarPayments", 5)]:
        budget = total_income * percentage / 100
        overspending_amount = float(overspending.get(category, 0))
        unexpected_expense_amount = float(unexpected_expenses.get(category, 0))
        remaining_budget = budget - overspending_amount - unexpected_expense_amount
        remaining_monthly_budget[category] = remaining_budget
    print("Thank you for completing the survey.")
    return (top_priority, financial_data, advice, actions, devotional, workout, budget_review, overspending, unexpected_expenses, total_income, remaining_monthly_budget)

# Example usage:
data = prompt_user()
print(data)
