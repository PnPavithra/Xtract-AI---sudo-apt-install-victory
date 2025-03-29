import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

# Load dataset
file_path = "Data/LibraryDataset.xls"  # Update path if needed
df = pd.read_excel(file_path)

# Convert 'Date' to datetime
df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

# Replace non-numeric values ("-") in 'Amount' with NaN and convert to numeric
df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce')

# Fill missing values in Amount with median
df['Amount'].fillna(df['Amount'].median(), inplace=True)

# Drop rows where essential columns are missing
df.dropna(subset=['Date', 'Title', 'Author', 'Department'], inplace=True)

# Set dark background theme
plt.style.use('dark_background')

# ---- Transactions Over Time ----
transactions_over_time = df['Date'].dt.date.value_counts().sort_index()
transactions_over_time.to_csv("SavedData/transactions_over_time.csv", header=['y'], index_label='x')

plt.figure(figsize=(14, 7))
transactions_over_time.plot(kind='line', color='cyan', marker='o', linewidth=2)
plt.xlabel('Date')
plt.ylabel('Transactions')
plt.title('Transactions Over Time')
plt.xticks(rotation=45)
plt.grid(alpha=0.3)
plt.tight_layout()
plt.show()

# ---- Top 10 Departments ----
top_departments = df['Department'].value_counts().nlargest(10)
top_departments.to_csv("SavedData/top_departments.csv", header=['y'], index_label='x')

plt.figure(figsize=(14, 8))
sns.barplot(x=top_departments.index, y=top_departments.values, palette='coolwarm')
plt.xlabel('Department')
plt.ylabel('Total Checkouts')
plt.title('Top 10 Departments Borrowing Books')
plt.xticks(rotation=30, ha='right')
plt.tight_layout()
plt.show()

# ---- Category-Wise Book Popularity ----
category_counts = df['Category'].value_counts()
category_counts.to_csv("category_counts.csv", header=['y'], index_label='x')

plt.figure(figsize=(10, 10))
plt.pie(category_counts, labels=category_counts.index, autopct='%1.1f%%', colors=sns.color_palette('Set2'))
plt.title('Category-Wise Book Borrowing')
plt.tight_layout()
plt.show()

# ---- Most Borrowed Books ----
top_books = df['Title'].value_counts().nlargest(10)
top_books.to_csv("SavedData/most_borrowed_books.csv", header=['y'], index_label='x')

plt.figure(figsize=(14, 8))
sns.barplot(x=top_books.values, y=top_books.index, palette='plasma')
plt.xlabel('Total Borrowed')
plt.ylabel('Book Title')
plt.title('Most Borrowed Books')
plt.tight_layout()
plt.show()

# ---- Most Borrowed Authors ----
top_authors = df['Author'].value_counts().nlargest(10)
top_authors.to_csv("SavedData/most_borrowed_authors.csv", header=['y'], index_label='x')

plt.figure(figsize=(14, 8))
sns.barplot(x=top_authors.values, y=top_authors.index, palette='viridis')
plt.xlabel('Total Borrowed')
plt.ylabel('Author Name')
plt.title('Most Borrowed Authors')
plt.tight_layout()
plt.show()

# ---- Least Borrowed Books ----
low_borrowed_books = df['Title'].value_counts().nsmallest(10)
low_borrowed_books.to_csv("SavedData/least_borrowed_books.csv", header=['y'], index_label='x')

plt.figure(figsize=(14, 8))
sns.barplot(x=low_borrowed_books.values, y=low_borrowed_books.index, palette='magma')
plt.xlabel('Total Borrowed')
plt.ylabel('Book Title')
plt.title('Least Borrowed Books (Underutilized)')
plt.tight_layout()
plt.show()

# ---- Expensive vs. Cheap Books ----
df['Price_Category'] = pd.qcut(df['Amount'], q=[0, 0.33, 0.66, 1], labels=['Cheap', 'Mid-Range', 'Expensive'])
price_distribution = df['Price_Category'].value_counts()
price_distribution.to_csv("SavedData/price_distribution.csv", header=['y'], index_label='x')

plt.figure(figsize=(10, 7))
sns.countplot(x=df['Price_Category'], palette=['green', 'orange', 'red'])
plt.xlabel('Book Price Category')
plt.ylabel('Count')
plt.title('Expensive vs. Cheap Books Distribution')
plt.grid(alpha=0.3)
plt.tight_layout()
plt.show()
