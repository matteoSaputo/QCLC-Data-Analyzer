import pandas as pd
import matplotlib.pyplot as plt

# Load the check-in form
check_in_form = pd.read_excel('course_validation.xlsx')

# Get the value counts for reasons
reason_counts = check_in_form['i am here to:'].value_counts()

# Calculate the total count
total_count = reason_counts.sum()

# Create the bar plot
fig, ax = plt.subplots(figsize=(12, 8))  # Adjust figure size as needed
bars = ax.bar(reason_counts.index, reason_counts.values, color='lightgreen')

# Annotate each bar with the percentage
for bar, value in zip(bars, reason_counts.values):
    percentage = (value / total_count) * 100
    ax.text(
        bar.get_x() + bar.get_width() / 2,  # X-coordinate: Center of the bar
        bar.get_height() + 1,              # Y-coordinate: Slightly above the bar
        f'{percentage:.1f}%',              # Label: Percentage with 1 decimal place
        ha='center',                       # Align horizontally to center
        fontsize=12                        # Font size for the annotations
    )

# Customize the chart
plt.title('Reasons for Visiting', fontsize=16)
plt.ylabel('Count', fontsize=14)
plt.xlabel('Reason', fontsize=14)

# Rotate x-axis labels for better readability, if needed
plt.xticks(fontsize=12, rotation=45, ha='right')
plt.yticks(fontsize=12)

# Save the chart as a PNG file
plt.tight_layout()  # Ensure everything fits within the figure
plt.savefig('Relevant_graphs/Reasons_for_Visiting.png')  # Save to the folder
plt.show() # Display the chart for debugging purposes (if using vscode, close the figure to finish running the script)
plt.close()  # Close the figure to free memory

print("Chart saved to Relevant_graphs/Reasons_for_Visiting.png")