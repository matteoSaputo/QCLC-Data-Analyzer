import pandas as pd
import matplotlib.pyplot as plt

# Load the check-in form and hours worked data
check_in_form = pd.read_excel('course_validation.xlsx')

check_in_form['start time'] = pd.to_datetime(check_in_form['start time'])
check_in_form['date'] = check_in_form['start time'].dt.date
check_in_form['hour'] = check_in_form['start time'].dt.hour

hourly_counts = check_in_form['hour'].value_counts().sort_index()

# Plot busy hours
hourly_counts.plot(kind='bar', title='Busy Times', figsize=(10, 6))
plt.ylabel('Count')
plt.xlabel('Hour')

# Save the chart
plt.tight_layout()
plt.savefig('Relevant_graphs/Busy_times.png')  # Save to the folder
plt.show() # Display the chart for debugging purposes
plt.close()  # Close the figure to free memory