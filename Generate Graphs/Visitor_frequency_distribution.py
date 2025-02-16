import pandas as pd
import matplotlib.pyplot as plt

# Load the check-in form
check_in_form = pd.read_excel('course_validation.xlsx')

# Count the number of visits per student
visit_counts = check_in_form['student id'].value_counts()

# Calculate the average number of visits per student
average_visits = visit_counts.mean()

# Create the histogram of visitation frequency
fig, ax = plt.subplots(figsize=(10, 6))  # Adjust figure size as needed
visit_counts.plot(kind='hist', bins=50, title='Distribution of Visit Frequency', ax=ax, color='skyblue')

# Add labels
plt.xlabel('Number of Visits', fontsize=14)
plt.ylabel('Number of Students', fontsize=14)

# Add the average visits as a text box
text_box_content = f'Average Visits per Student: {average_visits:.2f}'
ax.text(
    0.95, 0.95,  # Position: top-right corner
    text_box_content,
    transform=ax.transAxes,
    fontsize=12,
    bbox=dict(boxstyle="round,pad=0.3", edgecolor='black', facecolor='white'),
    ha='right'
)

# Save the chart as a PNG file
plt.tight_layout()  # Ensure everything fits within the figure
plt.savefig('Relevant_graphs/Visitor_Frequency_Distribution.png')  # Save to the folder
plt.show() # Display the chart for debugging purposes
plt.close()  # Close the figure to free memory

print("Visitor Frequency chart saved to Relevant_graphs/Visitor_Frequency_Distribution.png")