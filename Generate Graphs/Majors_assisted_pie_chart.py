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

# Create the pie chart
fig, ax = plt.subplots(figsize=(20, 10))
wedges, texts, autotexts = ax.pie(
    filtered_majors,
    labels=filtered_majors.index,
    autopct='%1.1f%%',
    startangle=90,
    textprops=dict(fontsize=15)
)

# Dynamically adjust the font size based on the percentage
for text, autotext, value in zip(texts, autotexts, filtered_majors):
    percentage = (value / total_count) * 100
    if percentage < 5:
        text.set_fontsize(8)
        autotext.set_fontsize(8)
    elif percentage < 10:
        text.set_fontsize(12)
        autotext.set_fontsize(12)
    else:
        text.set_fontsize(15)
        autotext.set_fontsize(15)

# Customize the chart
ax.set_title('Majors Assisted With', fontsize=20)
plt.ylabel('')

# Save the pie chart
plt.tight_layout()
plt.savefig('Relevant_graphs/Majors_Assisted_Pie_Chart.png')
plt.show()
plt.close()

print("Pie chart saved to Relevant_graphs/Majors_Assisted_Pie_Chart.png")
