import pandas as pd
import matplotlib.pyplot as plt

# Load the check-in form
check_in_form = pd.read_excel('course_validation.xlsx')

# Count the number of visits per student ID
visit_counts = check_in_form['student id'].value_counts()

# Get the top 30 frequent visitors
top_30_ids = visit_counts.head(30)

# Plot the bar chart
fig, ax = plt.subplots(figsize=(12, 8))
top_30_ids.plot(kind='bar', color='skyblue', ax=ax)

# Customize the chart
plt.title('Top 30 Frequent Visitors', fontsize=16)
plt.ylabel('Number of Visits', fontsize=14)
plt.xlabel('Student ID', fontsize=14)
plt.xticks(rotation=45, fontsize=10, ha='right')  # Rotate x-axis labels for better visibility

# Save the chart as a PNG file
plt.tight_layout()  # Ensure everything fits
plt.savefig('Relevant_graphs/Top_30_Frequent_Visitors.png')  # Save to the folder
plt.show() # Display the chart for debugging purposes
plt.close()  # Close the figure to free memory

print("Top 30 Frequent Visitors chart saved to Relevant_graphs/Top_30_Frequent_Visitors.png")