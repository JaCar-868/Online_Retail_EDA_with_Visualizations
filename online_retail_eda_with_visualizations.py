
# Import necessary libraries
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load the data
file_path = '/mnt/data/Online Retail.xlsx'
df = pd.read_excel(file_path)

# Initial exploration of the data
# Display the first few rows of the dataframe
print("First few rows of the dataframe:")
print(df.head())

# Display the summary statistics of the dataframe
print("\nSummary statistics of the dataframe:")
print(df.describe())

# Display the information about the dataframe
print("\nInformation about the dataframe:")
print(df.info())

# Visualize the data
# Plot the distribution of quantity
plt.figure(figsize=(10, 6))
sns.histplot(df['Quantity'], bins=50, kde=True)
plt.title('Distribution of Quantity')
plt.xlabel('Quantity')
plt.ylabel('Frequency')
plt.show()

# Plot the distribution of unit price
plt.figure(figsize=(10, 6))
sns.histplot(df['UnitPrice'], bins=50, kde=True)
plt.title('Distribution of Unit Price')
plt.xlabel('Unit Price')
plt.ylabel('Frequency')
plt.show()

# Clean the data
# Check for missing values
print("\nMissing values in the dataframe:")
print(df.isnull().sum())

# Drop rows with missing CustomerID
df_cleaned = df.dropna(subset=['CustomerID'])

# Remove rows with negative or zero quantities
df_cleaned = df_cleaned[df_cleaned['Quantity'] > 0]

# Remove rows with negative or zero unit price
df_cleaned = df_cleaned[df_cleaned['UnitPrice'] > 0]

# Re-explore the cleaned data
print("\nFirst few rows of the cleaned dataframe:")
print(df_cleaned.head())

print("\nSummary statistics of the cleaned dataframe:")
print(df_cleaned.describe())

# Perform further analysis
# Top 10 countries by number of transactions
top_countries = df_cleaned['Country'].value_counts().head(10)
plt.figure(figsize=(12, 8))
sns.barplot(x=top_countries.values, y=top_countries.index, palette='viridis')
plt.title('Top 10 Countries by Number of Transactions')
plt.xlabel('Number of Transactions')
plt.ylabel('Country')
plt.show()

# Monthly sales analysis
df_cleaned['InvoiceDate'] = pd.to_datetime(df_cleaned['InvoiceDate'])
df_cleaned.set_index('InvoiceDate', inplace=True)
monthly_sales = df_cleaned['Quantity'].resample('M').sum()
plt.figure(figsize=(14, 7))
monthly_sales.plot()
plt.title('Monthly Sales')
plt.xlabel('Month')
plt.ylabel('Quantity Sold')
plt.show()

# Additional Visualizations
# Sales distribution by country
df_cleaned['TotalSales'] = df_cleaned['Quantity'] * df_cleaned['UnitPrice']
country_sales = df_cleaned.groupby('Country')['TotalSales'].sum().sort_values(ascending=False).head(10)
plt.figure(figsize=(12, 8))
sns.barplot(x=country_sales.values, y=country_sales.index, palette='viridis')
plt.title('Top 10 Countries by Total Sales')
plt.xlabel('Total Sales')
plt.ylabel('Country')
plt.show()

# Top 10 most sold products
top_products = df_cleaned.groupby('Description')['Quantity'].sum().sort_values(ascending=False).head(10)
plt.figure(figsize=(12, 8))
sns.barplot(x=top_products.values, y=top_products.index, palette='viridis')
plt.title('Top 10 Most Sold Products')
plt.xlabel('Quantity Sold')
plt.ylabel('Product')
plt.show()

# Monthly revenue analysis
monthly_revenue = df_cleaned['TotalSales'].resample('M').sum()
plt.figure(figsize=(14, 7))
monthly_revenue.plot()
plt.title('Monthly Revenue')
plt.xlabel('Month')
plt.ylabel('Revenue')
plt.show()

# Customer distribution by country
customer_dist = df_cleaned['Country'].value_counts().head(10)
plt.figure(figsize=(12, 8))
sns.barplot(x=customer_dist.values, y=customer_dist.index, palette='viridis')
plt.title('Top 10 Countries by Number of Customers')
plt.xlabel('Number of Customers')
plt.ylabel('Country')
plt.show()

# Save the cleaned dataframe to a new Excel file
df_cleaned.to_excel('/mnt/data/Cleaned_Online_Retail.xlsx', index=False)
