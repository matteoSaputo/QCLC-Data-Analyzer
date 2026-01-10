import pandas as pd
import matplotlib.pyplot as plt

# Load the check-in form
check_in_form = pd.read_excel('course_validation.xlsx')

# Parse datetime
check_in_form['start time'] = pd.to_datetime(check_in_form['timestamp'])
check_in_form['date'] = check_in_form['start time'].dt.date
check_in_form['weekday'] = check_in_form['start time'].dt.weekday  # 0=Monday, 6=Sunday

# Exclude weekends
check_in_form = check_in_form[check_in_form['weekday'] < 5]  # Keep Mon-Fri

# Define a list of holidays
holidays = [
    '2025-01-01', '2025-01-20', '2025-02-17', '2025-05-26',
    '2025-07-04', '2025-09-01', '2025-11-11', '2025-11-27', '2025-12-25'
]
holidays = pd.to_datetime(holidays).date

check_in_form = check_in_form[check_in_form['i am here to'].str.lower().str.contains('tutor', na=False)]

# Exclude holidays
check_in_form = check_in_form[~check_in_form['date'].isin(holidays)]

# Count occurrences by date
date_counts = check_in_form['date'].value_counts().sort_index()

# Plot
date_counts.plot(kind='line', title='Tutoring Activity Volume over the Semester', figsize=(10, 6))
plt.ylabel('Visitations')
plt.xlabel('Date')

# Save
plt.tight_layout()
plt.savefig('./QCLC-Data-Analyzer/Relevant_graphs/Busy_tutoring_dates.png')
plt.show()
plt.close()
