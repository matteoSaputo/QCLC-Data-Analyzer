import pandas as pd
import matplotlib.pyplot as plt

# Load the check-in form and hours worked data
check_in_form = pd.read_excel('course_validation.xlsx')
hours_worked = pd.read_excel('tutor_hours_worked.xlsx')

# Standardize the tutor names in both DataFrames
check_in_form['who assisted?'] = check_in_form['who assisted?'].str.strip().str.lower()
hours_worked['Tutor Name'] = hours_worked['Tutor Name'].str.strip().str.lower()

# Ensure 'start time' is in datetime format
check_in_form['start time'] = pd.to_datetime(check_in_form['start time'])

# Calculate the total number of students helped by each tutor
total_students_helped = check_in_form.groupby('who assisted?').size().reset_index(name='students_helped')

# Merge the total students helped with the hours worked
tutor_data = total_students_helped.merge(
    hours_worked,
    left_on='who assisted?',
    right_on='Tutor Name',
    how='inner'
)

# Debug: Check if the merge worked
if tutor_data.empty:
    print("No matching tutors found between the check-in form and the hours worked data.")
else:
    # Calculate the true hourly average
    tutor_data['students_per_hour'] = tutor_data['students_helped'] / tutor_data['Hours Worked']
    # Sort the data by 'students_per_hour' in descending order
    tutor_data = tutor_data.sort_values(by='students_per_hour', ascending=False, key=lambda col: col.fillna(0))
    
    # Plot the sorted bar chart
    fig, ax = plt.subplots(figsize=(20, 10))  # Adjust figure size
    bars = ax.bar(tutor_data['Tutor Name'], tutor_data['students_per_hour'], color='skyblue')
    
    # Annotate each bar with the average
    for bar, avg in zip(bars, tutor_data['students_per_hour']):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() + 0.01,  # Adjust position: closer and below the top of the bar
            f'{avg:.2f}',  # Label: Average with 2 decimal places
            ha='center',
            va='bottom',  # Align with bottom of the label
            fontsize=12
        )
    
    # Customize the chart
    plt.title('Hourly Average of Students Helped per Tutor', fontsize=20)
    plt.ylabel('Average Students Helped per Hour', fontsize=16)
    plt.xlabel('Tutor', fontsize=16)
    
    # Rotate x-axis labels for better readability
    plt.xticks(rotation=45, ha='right', fontsize=12)
    plt.yticks(fontsize=12)
    
    # Adjust layout to ensure labels fit
    plt.tight_layout()
    
    # Save the chart
    plt.savefig('Relevant_graphs/Tutor_Assistance_Average_Hourly.png')  # Save the chart
    plt.show()
