import matplotlib.pyplot as plt
from datetime import datetime
import copy
import sys
import os

colors = ['b', 'r', 'g', 'c', 'm', 'y', 'k']


class Transaction:
    def __init__(self, date, name, subcategory, category, amount):
        self.date = datetime.strptime(date, "%m-%d-%Y")
        self.name = name
        self.subcategory = subcategory
        self.category = category
        self.amount = float(amount.replace('$', ''))

    def month_id(self):
        return self.date.strftime("%b-%Y")


def read_file(fname):
    with open(fname) as f:
        lines = f.readlines()

    transactions = []
    for line in lines:
        elements = line.strip().split(',')

        # Skip uncategorized
        if elements[3] == "Uncategorized":
            continue

        new_transaction = Transaction(
            elements[0], elements[1], elements[2], elements[3], elements[4])
        transactions.append(new_transaction)

    return transactions


def get_unique_categories(transactions):
    unique_categories = []
    for transaction in transactions:
        if transaction.category not in unique_categories:
            unique_categories.append(transaction.category)

    return unique_categories


def month_sorter(e):
    monthYear = e.split('-')
    month = datetime.strptime(monthYear[0], "%b").month
    year = int(monthYear[1])

    # Returns int such as '202101' for Jan 2021
    return year * 100 + month


def get_unique_months(transactions):
    unique_months = []
    for transaction in transactions:
        if transaction.month_id() not in unique_months:
            unique_months.append(transaction.month_id())

    # Sort months
    unique_months.sort(key=month_sorter)

    return unique_months


def group_by_month_category(transactions, unique_months):
    # Get category_month_map
    category_month_map = {}
    for transaction in transactions:
        month_id = transaction.month_id()
        category = transaction.category
        amount = transaction.amount

        if category not in category_month_map:
            category_month_map[category] = {}

        if month_id not in category_month_map[category]:
            category_month_map[category][month_id] = 0

        category_month_map[category][month_id] += amount

    # Reduce to just category_map
    # Each map points to a list of elements based where each element is one time frame
    category_map = {}
    for category in category_month_map:
        amount_list_per_category = []

        for month in unique_months:
            if month not in category_month_map[category]:
                amount = 0
            else:
                amount = category_month_map[category][month]
            amount_list_per_category.append(amount)

        category_map[category] = amount_list_per_category

    return category_map


def graph_by_month_category(category_map, unique_months, fname):
    fig = plt.figure()
    ax = fig.add_axes([0.15, 0.2, 0.55, 0.7])

    bottom = [0] * len(unique_months)
    for idx, category in enumerate(category_map):
        ax.bar(unique_months, category_map[category],
               0.4, color=colors[idx], bottom=bottom)

        for val_idx, val in enumerate(category_map[category]):
            bottom[val_idx] += val

    ax.legend(labels=category_map.keys(), loc=(1.04, 0))

    ax.set_ylabel("Spend")
    ax.set_xlabel("Time")
    plt.xticks(rotation=90)

    plt.savefig(fname)


def remove_all_keys_except(category_map, str):
    new_category_map = copy.deepcopy(category_map)
    for key in list(category_map.keys()):
        if key != str:
            new_category_map.pop(key, None)
    return new_category_map


def main():
    fname = sys.argv[1]
    transactions = read_file(fname)
    unique_months = get_unique_months(transactions)
    category_map = group_by_month_category(
        transactions, unique_months)

    if not os.path.exists("images"):
        os.makedirs("images")

    for key in category_map:
        new_category_map = remove_all_keys_except(category_map, key)
        graph_by_month_category(
            new_category_map, unique_months, f"images/{key.lower()}.jpg")

    graph_by_month_category(category_map, unique_months, "images/all.jpg")


main()
