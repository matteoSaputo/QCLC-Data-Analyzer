import pandas as pd
import matplotlib.pyplot as plt

# Load the check-in form
check_in_form = pd.read_excel('course_validation.xlsx')

# Extract major from course names (assumes major is the prefix before the space)
check_in_form['Major'] = check_in_form['what course do you need assistance with?'].str.split().str[0]

# Get the value counts of majors
major_counts = check_in_form['Major'].value_counts()

# Define the threshold percentage
threshold_percentage = 0.75  # Combine majors contributing less than 0.75% of the total

# Calculate the total count
total_count = major_counts.sum()

# Separate major categories and group smaller ones into "Other"
filtered_majors = major_counts[major_counts / total_count * 100 >= threshold_percentage]
other_count = major_counts[major_counts / total_count * 100 < threshold_percentage].sum()

# Add "Other" to the filtered majors if applicable
if other_count > 0:
    filtered_majors['Other'] = other_count

# Create the bar graph for majors
fig, ax = plt.subplots(figsize=(20, 10))
bars = ax.bar(filtered_majors.index, filtered_majors.values, color='skyblue')

# Annotate each bar with the percentage
for bar, value in zip(bars, filtered_majors):
    percentage = (value / total_count) * 100
    ax.text(
        bar.get_x() + bar.get_width() / 2,
        bar.get_height() + 1,
        f'{percentage:.1f}%',
        ha='center',
        fontsize=12
    )

# Customize the bar graph
ax.set_title('Majors Assisted With (Bar Graph)', fontsize=20)
ax.set_xlabel('Major', fontsize=16)
ax.set_ylabel('Number of Students', fontsize=16)
plt.xticks(rotation=45, fontsize=12)  # Rotate x-axis labels for readability
plt.yticks(fontsize=12)

# Save the bar graph
plt.tight_layout()
plt.savefig('Relevant_graphs/Majors_Assisted_Bar_Graph.png')  # Save to the folder
plt.show()  # Display the chart for debugging purposes
plt.close()  # Close the figure to free memory

print("Bar graph saved to Relevant_graphs/Majors_Assisted_Bar_Graph.png")