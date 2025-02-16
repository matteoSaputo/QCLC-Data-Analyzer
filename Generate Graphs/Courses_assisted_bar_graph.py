import pandas as pd
import matplotlib.pyplot as plt

# Load the check-in form
check_in_form = pd.read_excel('course_validation.xlsx')

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

# Create the bar graph
fig, ax = plt.subplots(figsize=(20, 10))
bars = ax.bar(filtered_courses.index, filtered_courses.values, color='skyblue')

# Annotate each bar with the percentage
for bar, value in zip(bars, filtered_courses):
    percentage = (value / total_count) * 100
    ax.text(
        bar.get_x() + bar.get_width() / 2,
        bar.get_height() + 1,
        f'{percentage:.1f}%',
        ha='center',
        fontsize=12
    )

# Customize the bar graph
ax.set_title('Courses Assisted With (Bar Graph)', fontsize=20)
ax.set_xlabel('Course', fontsize=16)
ax.set_ylabel('Number of Students', fontsize=16)
plt.xticks(rotation=45, fontsize=12)  # Rotate x-axis labels for readability
plt.yticks(fontsize=12)

# Save the bar graph
plt.tight_layout()
plt.savefig('Relevant_graphs/Courses_Assisted_Bar_Graph.png')  # Save to the folder
plt.show() # Display the chart for debugging purposes
plt.close()  # Close the figure to free memory

print("Bar graph saved to Relevant_graphs/Courses_Assisted_Bar_Graph.png")