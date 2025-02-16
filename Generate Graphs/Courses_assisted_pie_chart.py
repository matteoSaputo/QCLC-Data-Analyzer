import pandas as pd
import matplotlib.pyplot as plt

# Load the check-in form
check_in_form = pd.read_excel('course_validation.xlsx')

import matplotlib.pyplot as plt

# Get the value counts of courses
course_counts = check_in_form['what course do you need assistance with?'].value_counts()

# Define the threshold percentage
threshold_percentage = 0.75  # Combine courses contributing less than 0.75% of the total

# Calculate the total count
total_count = course_counts.sum()

# Separate major courses and group smaller ones into "Other"
filtered_courses = course_counts[course_counts / total_count * 100 >= threshold_percentage]
other_count = course_counts[course_counts / total_count * 100 < threshold_percentage].sum()

# Add "Other" to the filtered courses if applicable
if other_count > 0:
    filtered_courses['Other'] = other_count

# Create the pie chart
fig, ax = plt.subplots(figsize=(20, 10))
wedges, texts, autotexts = ax.pie(
    filtered_courses,
    labels=filtered_courses.index,
    autopct='%1.1f%%',
    startangle=90,
    textprops=dict(fontsize=15)
)

# Dynamically adjust the font size based on the percentage
for text, autotext, value in zip(texts, autotexts, filtered_courses):
    percentage = (value / total_count) * 100
    if percentage < 5:
        text.set_fontsize(8)  # Smaller font for less significant slices
        autotext.set_fontsize(8)
    elif percentage < 10:
        text.set_fontsize(12)
        autotext.set_fontsize(12)
    else:
        text.set_fontsize(15)  # Larger font for more significant slices
        autotext.set_fontsize(15)

# Customize the chart
ax.set_title('Courses Assisted With', fontsize=20)
plt.ylabel('')  # Remove the default y-axis label

# Save the pie chart
plt.tight_layout()
plt.savefig('Relevant_graphs/Courses_Assisted_Pie_Chart.png')  # Save to the folder
plt.show() # Display the chart for debugging purposes
plt.close()  # Close the figure to free memory

print("Pie chart saved to Relevant_graphs/Courses_Assisted_Pie_Chart.png")
