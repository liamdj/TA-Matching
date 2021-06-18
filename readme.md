# TA Matching

This project aims to automate assignments betwwen graduate student TAs and undergraduate courses within the Princeton Computer Science Department.

## Input

The matching algorithm takes input specified below. See sample input in /test.

#### Student
For each student,
- Previous courses that the student TAed
- Course that the student's advisor is teaching
- For each course, how well the student believes they would match based on course information previously solicited from professors
    - If the student indicates 'Unqualified', the student cannot be placed in that course.
    - Otherwise, the student selects from a linear scale from poor (1) to excellent (4).
- An optional favorite course

[Here](https://docs.google.com/forms/d/1JTzWBqIFEKQ8nNFOA7JMkzzfylRxmQox8tFRjcXZW8I/edit?usp=sharing) is a sample form to generate responses (download responses as csv).

### Professor
For each course,
- Optionally for any student, how well the professor believes they would match based on the previously solicited student's information
    - If the professor selects 'Distaster', the student cannnot be placed in that course.
    - Otherwise, the professors selects from a linear scale from poor (1) to excellent (4). 

[Here](https://docs.google.com/forms/d/1G0g8RtkN2ISPQXZh0tqsTUHNHaM_tPeNeTSPFoeCh48/edit?usp=sharing) is a sample form to generate responses (download responses as csv).

### Course info
For each course,
- 'Target TAs'
    - This number is the maximum and prefered number of TAs in the course.
- Optional 'Fill Weight'
    - Indicates how many weight is given to filling all target TA slots in the course for the cases when the number of students is less than the total target number of TAs. Default is 1.

### Fixed
Optionally specify students-course pairs that will be automatically assigned, so will not need to be considered when making the other matchings.

### Adjusted
Optionally specify student-course pairs whose weight will be modified by 'Weight' beforing running the matching.

## Output

The algorithm finds the assignments producing the highest total weight and writes to a csv file.

## Example Usage

In the project directory root, running
```
python matching.py ./test/Student.csv ./test/Prof.csv ./test/Course.csv --fixed ./test/Fixed.csv --adjusted ./test/Adjusted.csv --matchings ./test/Matchings.csv
```
will preform the algorithm on the test inputs and save the matchings at `./test/Matchings.csv`.
