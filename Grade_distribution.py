import pandas as pd
import matplotlib.pyplot as plt

# Possible Grades
grades = ['A+', 'A', 'A-', 'B+', 'B', 'B-', 'C+', 'C', 'C-', 'D+', 'D', 'F', 'W', 'WU', 'NC', 'INC']

# Define a mapping for grades to numeric values
grade_to_numeric = {
    'A+': 4.3, 'A': 4.0, 'A-': 3.7,
    'B+': 3.3, 'B': 3.0, 'B-': 2.7,
    'C+': 2.3, 'C': 2.0, 'C-': 1.7,
    'D+': 1.3, 'D': 1.0, 'F': 0.0,
    'W': None, 'WU': None, 'NC': None, 'INC': None
}

# Load cleaned check-in form and roster data
check_in_form = pd.read_excel('course_validation.xlsx')
roster = pd.read_excel('combined_roster.xlsx')

# Map grades to numeric values in the roster
roster['grade_numeric'] = roster['grade'].map(grade_to_numeric)

# Standardize column formats
check_in_form.columns = check_in_form.columns.str.lower().str.strip()
roster.columns = roster.columns.str.lower().str.strip()
check_in_form['student id'] = check_in_form['student id'].astype(str).str.strip()
roster['id'] = roster['id'].astype(str).str.strip()
roster['catalog'] = roster['catalog'].astype(str).str.strip()

# Get all unique courses from the check-in form
unique_courses = check_in_form['what course do you need assistance with?'].dropna().unique()

# Loop through each course and create a graph
for course in unique_courses:
    # Extract subject and catalog from the course name
    parts = course.split()  # Splits on any whitespace, ignoring multiples
    if len(parts) < 2:
        print(f"Invalid course format: {course}")
        continue

    # Rejoin everything after the first token into the catalog
    subject = parts[0]
    catalog = " ".join(parts[1:])  # Handles multiple tokens after the subject


    # Filter data for the course
    check_in_course = check_in_form[check_in_form['what course do you need assistance with?'] == course]
    roster_course = roster[(roster['subject'] == subject) & (roster['catalog'] == catalog)]

    # Add a tutoring flag to the roster
    roster_course['tutored'] = roster_course['id'].isin(check_in_course['student id'])

    # Group and count grades
    lc_visitor_counts = roster_course[roster_course['tutored']].groupby('grade')['id'].count().reindex(grades, fill_value=0)
    general_population_counts = roster_course.groupby('grade')['id'].count().reindex(grades, fill_value=0)

    # Calculate percentages
    lc_visitor_percentages = (lc_visitor_counts / lc_visitor_counts.sum()) * 100
    general_population_percentages = (general_population_counts / general_population_counts.sum()) * 100

    # Calculate average grades
    lc_average_grade = roster_course[roster_course['tutored']]['grade_numeric'].mean()
    general_average_grade = roster_course['grade_numeric'].mean()
    
    # Combine data into a single DataFrame
    grade_comparison = pd.DataFrame({
        'LC Visitors (%)': lc_visitor_percentages,
        'General Population (%)': general_population_percentages
    }).fillna(0)

    # Plot the bar graph
    fig, ax = plt.subplots(figsize=(12, 8))
    x = range(len(grades))  # X-axis positions
    bar_width = 0.4

    # Plot LC Visitors bars
    lc_bars = ax.bar(
        [pos - bar_width / 2 for pos in x],
        grade_comparison['LC Visitors (%)'],
        bar_width,
        label='LC Visitors',
        color='skyblue'
    )

    # Plot General Population bars
    gp_bars = ax.bar(
        [pos + bar_width / 2 for pos in x],
        grade_comparison['General Population (%)'],
        bar_width,
        label='General Population',
        color='orange'
    )

    # Add percentages above the bars
    for bar in lc_bars:
        height = bar.get_height()
        if height > 0:  # Only annotate non-zero bars
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                height + 0.5,
                f'{height:.1f}%',
                ha='center',
                fontsize=10
            )

    for bar in gp_bars:
        height = bar.get_height()
        if height > 0:  # Only annotate non-zero bars
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                height + 0.5,
                f'{height:.1f}%',
                ha='center',
                fontsize=10
            )

    # Calculate totals
    total_tutored = roster_course['tutored'].sum()
    total_students = len(roster_course)

    # Add a text box with totals and average grades
    text_box_content = (
        f"Total Tutored: {total_tutored}\n"
        f"Total Students: {total_students}\n"
        f"Avg Grade (Tutored): {lc_average_grade:.2f}\n"
        f"Avg Grade (General): {general_average_grade:.2f}"
    )
    ax.text(
        1.05, 0.95,
        text_box_content,
        transform=ax.transAxes,
        fontsize=12,
        bbox=dict(boxstyle="round,pad=0.3", edgecolor='black', facecolor='white')
    )

    # Customize the plot
    ax.set_xticks(x)
    ax.set_xticklabels(grades, fontsize=12)
    ax.set_title(f'Grade Distribution for {course}', fontsize=16)
    ax.set_xlabel('Grades', fontsize=14)
    ax.set_ylabel('Percentage of Students (%)', fontsize=14)
    ax.legend(fontsize=12)
    plt.tight_layout()

    # Save the plot to a file
    plt.savefig(f'Grade_Distributions/Grade_Distribution_{course}.png')  # Saves each graph as a PNG file
    plt.close()  # Close the plot to save memory

print("Graphs for all courses have been generated and saved.")