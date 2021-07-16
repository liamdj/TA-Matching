# TA Matching

This project aims to automate assignments betwwen graduate student TAs and undergraduate courses within the Princeton Computer Science Department.

## Input

The matching algorithm takes input specified below. See sample input in /test.

### Student Preferences
For each student,
- For each of five categorical rankings, which courses the student ranks the matching based on course information previously solicited from professors
    - If the student indicates 'Unqualified', the student cannot be placed in that course.
    - Otherwise, the student selects from a linear scale from poor (1) to excellent (4).
- An optional favorite course

[Here](https://docs.google.com/forms/d/1qp0XepvZQzHhLZe5gHkh6pI7_Enh5cWPA8GeWvgvAkM/edit?usp=sharing) is a sample form to generate responses (download responses as csv).

### Professor Preferences
For each course,
- For each of five categorical rankings, optionally which students the professor ranks the matching based on previously solicited student's information
    - If the professor indicates 'Distaster', the student cannnot be placed in that course.
    - Otherwise, the professors selects from a linear scale from poor (1) to excellent (4). 

[Here](https://docs.google.com/forms/d/1OcUSU5a2dRylvHlTwJ-H---dSMm4V41jWGp0LV9UfCg/edit?usp=sharing) is a sample form to generate responses (download responses as csv).

### Student Information
Optionally for each student, 
- 'Previous Courses' that the student TAed
- 'Advisor's Course' that the student's advisor is teaching
- 'Assign Weight' describing how desirable it is to find a match for the student
    - Only matters when there are more students that TA slots
    - Default is 0

### Course Information
For each course,
- 'TA Capacity' describing the maximum number of TAs in the course.
- Optional 'Fill Weight' describing how desirable it is to fill the TA slots in the course 
    - Only matters whene there are fewer students than TA slots
    - Default is 1

### Fixed
Optionally specify students-course pairs that will be automatically assigned, so will not need to be considered when making the other matchings.

### Adjusted
Optionally specify student-course pairs whose weight will be modified by 'Weight' before running the matching.

## Output

### Matching
The algorithm constructs a bipartite graph containing a source node, sink node, and a node for each student and each course TA slot. The assignments producing the highest total weight is written to a csv file.

## Example Usage

In the project directory root, running
```
python matching.py ./test/Student.csv ./test/Prof.csv ./test/Course.csv --fixed ./test/Fixed.csv --adjusted ./test/Adjusted.csv --matchings ./test/Matchings.csv
```
will preform the algorithm on the test inputs and save the matchings at `./test/Matchings.csv`.
