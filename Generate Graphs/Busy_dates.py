import pandas as pd
import matplotlib.pyplot as plt

# Load the check-in form and hours worked data
check_in_form = pd.read_excel('course_validation.xlsx')

check_in_form['start time'] = pd.to_datetime(check_in_form['start time'])
check_in_form['date'] = check_in_form['start time'].dt.date
check_in_form['hour'] = check_in_form['start time'].dt.hour

date_counts = check_in_form['date'].value_counts().sort_index()

# Plot busy dates
date_counts.plot(kind='line', title='Busy Dates', figsize=(10, 6))
plt.ylabel('Count')
plt.xlabel('Date')

# Save the chart
plt.tight_layout()
plt.savefig('Relevant_graphs/Busy_dates.png')  # Save to the folder
plt.show() # Display the chart for debugging purposes
plt.close()  # Close the figure to free memory