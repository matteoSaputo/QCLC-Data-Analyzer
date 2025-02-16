import pandas as pd
import matplotlib.pyplot as plt

# Load the check-in form
check_in_form = pd.read_excel('course_validation.xlsx')

# Get the value counts for appointments
appointment_counts = check_in_form['do you have an appointment?'].value_counts()

# Calculate the total counts
total_count = appointment_counts.sum()

# Create the bar plot
fig, ax = plt.subplots(figsize=(10, 6))  # Adjust figure size as needed
bars = ax.bar(appointment_counts.index, appointment_counts.values, color='skyblue')

# Annotate each bar with the percentage
for bar, value in zip(bars, appointment_counts.values):
    percentage = (value / total_count) * 100
    ax.text(
        bar.get_x() + bar.get_width() / 2,  # X-coordinate: Center of the bar
        bar.get_height() + 1,              # Y-coordinate: Slightly above the bar
        f'{percentage:.1f}%',              # Label: Percentage with 1 decimal place
        ha='center',                       # Align horizontally to center
        fontsize=12                        # Font size for the annotations
    )

# Customize the chart
plt.title('Walk-Ins vs Appointments', fontsize=16)
plt.ylabel('Count', fontsize=14)
plt.xlabel('Type', fontsize=14)

# Rotate x-axis labels if needed (not likely for this chart)
plt.xticks(fontsize=12)
plt.yticks(fontsize=12)

# Save the chart as a PNG file
plt.tight_layout()  # Ensure everything fits within the figure
plt.savefig('Relevant_graphs/Walk_Ins_vs_Appointments.png')  # Save to the folder
plt.show() # Display the chart for debugging purposes (if using vscode, close the figure to finish running the script)
plt.close()  # Close the figure to free memory

print("Chart saved to Relevant_graphs/Walk_Ins_vs_Appointments.png")