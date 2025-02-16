import pandas as pd
import matplotlib.pyplot as plt

# Load the check-in form
check_in_form = pd.read_excel('course_validation.xlsx')

# Ensure 'start time' is in datetime format
check_in_form['start time'] = pd.to_datetime(check_in_form['start time'])

# Extract year and week number from 'start time'
check_in_form['week'] = check_in_form['start time'].dt.isocalendar().week
check_in_form['year'] = check_in_form['start time'].dt.year

# Group by tutor and week to count students helped per week
weekly_tutor_counts = check_in_form.groupby(['who assisted?', 'year', 'week']).size().reset_index(name='students_helped')

# Calculate the average number of students helped per week for each tutor
tutor_weekly_averages = weekly_tutor_counts.groupby('who assisted?')['students_helped'].mean()

# Sort by the average to show busiest tutors first
tutor_weekly_averages = tutor_weekly_averages.sort_values(ascending=False)

# Plot the bar chart
fig, ax = plt.subplots(figsize=(20, 10))  # Adjust figure size as needed
bars = ax.bar(tutor_weekly_averages.index, tutor_weekly_averages.values, color='skyblue')

# Annotate each bar with the average
for bar, avg in zip(bars, tutor_weekly_averages.values):
    ax.text(
        bar.get_x() + bar.get_width() / 2,
        bar.get_height() + 0.2,
        f'{avg:.1f}',  # Label: Average with 1 decimal place
        ha='center',
        fontsize=12
    )

# Customize the chart
plt.title('Average Number of Students Helped per Week by Tutor', fontsize=20)
plt.ylabel('Average Students Helped per Week', fontsize=16)
plt.xlabel('Tutor', fontsize=16)

# Rotate x-axis labels for better readability
plt.xticks(rotation=45, ha='right', fontsize=12)
plt.yticks(fontsize=12)

# Save the chart
plt.tight_layout()
plt.savefig('Relevant_graphs/Tutor_Assistance_Average_Weekly.png')  # Save the chart
plt.show() # Display the chart for debugging purposes
plt.close()  # Close the figure to free memory

print("Tutor Assistance chart saved to Relevant_graphs/Tutor_Assistance_Average_Weekly.png")