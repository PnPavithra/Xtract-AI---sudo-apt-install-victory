import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load Dataset
file_path = "Data/LibraryDataset.xls"  # Change if needed
df = pd.read_excel(file_path)

# Convert Date to datetime
df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

# Remove missing values
df.dropna(subset=["Date", "Title", "Department", "Transaction"], inplace=True)

# Set dark background
plt.style.use("dark_background")

### NEGATIVE METRICS VISUALIZATIONS ###

import os

def save_chart(fig, title):
    """ Saves figure as a high-resolution image in the current working directory """
    save_path = os.path.join(os.getcwd(), f"{title}.png")  # Saves in script's directory
    fig.savefig(save_path, dpi=300, bbox_inches="tight")
    print(f"Saved: {save_path}")


# 1. Underutilized Books (Least Borrowed)
least_borrowed_books = df["Title"].value_counts().nsmallest(10)
fig, ax = plt.subplots(figsize=(12, 6))
least_borrowed_books.plot(kind="barh", color="red", ax=ax)
ax.set_title("Top 10 Least Borrowed Books")
ax.set_xlabel("Number of Checkouts")
ax.set_ylabel("Book Title")
plt.xticks(color="white")
plt.yticks(color="white")
save_chart(fig, "least_borrowed_books")


# 2. Departments with Least Borrowing
least_borrowing_depts = df["Department"].value_counts().nsmallest(10)
fig, ax = plt.subplots(figsize=(12, 6))
least_borrowing_depts.plot(kind="barh", color="orange", ax=ax)
ax.set_title("Bottom 10 Departments by Book Borrowing")
ax.set_xlabel("Number of Transactions")
ax.set_ylabel("Department")
plt.xticks(color="white")
plt.yticks(color="white")
save_chart(fig, "least_borrowing_departments")


# 3. Books with Longest Checkout Duration (Potential Hoarding)
df["Days Checked Out"] = (df["Date"].max() - df["Date"]).dt.days
longest_held_books = df.groupby("Title")["Days Checked Out"].max().nlargest(10)
fig, ax = plt.subplots(figsize=(12, 6))
longest_held_books.plot(kind="barh", color="purple", ax=ax)
ax.set_title("Top 10 Books with Longest Checkout Duration")
ax.set_xlabel("Days Checked Out")
ax.set_ylabel("Book Title")
plt.xticks(color="white")
plt.yticks(color="white")
save_chart(fig, "longest_checked_out_books")


# 4. Most Unpopular Authors (Least Borrowed)
least_borrowed_authors = df["Author"].value_counts().nsmallest(10)
fig, ax = plt.subplots(figsize=(12, 6))
least_borrowed_authors.plot(kind="barh", color="yellow", ax=ax)
ax.set_title("Top 10 Least Borrowed Authors")
ax.set_xlabel("Number of Checkouts")
ax.set_ylabel("Author")
plt.xticks(color="white")
plt.yticks(color="white")
save_chart(fig, "least_borrowed_authors")


# 5. Branch Disparities (Books Staying in Home Branch)
branch_issues = df.groupby(["homebranch", "holdingbranch"]).size().unstack()
fig, ax = plt.subplots(figsize=(10, 6))
sns.heatmap(branch_issues, cmap="coolwarm", linewidths=1, ax=ax)
ax.set_title("Library Branch Disparities in Book Movement")
plt.xticks(color="white", rotation=45)
plt.yticks(color="white", rotation=0)
save_chart(fig, "branch_disparities")


# 6. Zero Transaction Books (if we had book inventory data)
# Assuming we had a list of all books and their checkouts:
# available_books = pd.read_excel("book_inventory.xls")
# zero_transaction_books = set(available_books["Title"]) - set(df["Title"])
# If you have this data, I can add the graph.

print("All negative metrics graphs saved successfully!")
