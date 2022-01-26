# TA Matching

This project aims to automate assignments betwwen graduate student TAs and undergraduate courses within the Princeton Computer Science Department.

## Requirements
To run the matching, python 3 and the following modules are required, which can be installed with 
```
pip install -r requirements.txt
``` 

## Input

The matching algorithm takes the input specified below. See sample input in `test`. Note that if the students in student preferences and student information do not match or the courses in professor preferences and course informition do not match, the program will terminate with a relevant error message. The file generate_random_input.py takes a course input file and creates the other three required input files.

### Student Data
For each student,
- NetID
- Name
- Weight: describes how desirable it is to find a match for the student (default 10)
- Previous: courses that the student previously TAed
- Advisors: courses that a student's advisor is teaching
- Favorite: best course matches
- Good: decent course matches
- Okay: satisfactory course matches

Unspecified courses are considered unsatisfactory matches and will never be assigned.

### Professor Data
For each course,
- Course: name
- Slots: maximum number of TAs in the course
- Weight: how desirable it is to fill slots in the course (default 10)
- Favorite: netids of students professor wants as TA
- Veto: netids of students the professor will not have as TA

Unspecifed students are considered neutral and may be matched.

### Fixed
Optionally specify students-course pairs that will be forced to be assigned.

### Adjusted
Optionally specify student-course pairs whose weight will be modified by "Weight" before running the matching.

## Parameters
The tunable parameters used to construct weights are:
- PREVIOUS_WEIGHT: units to increase matches where the student has previously TAed
- ADVISORS_WEIGHT: units to increase advisor-advizee matches
- STUDENT_PREF_WEIGHT: units per standard deviation in a student's rankings
- PROF_PREF_WEIGHT: units per standard deviation in a professor's rankings

## Output

### Matching
The algorithm constructs a bipartite graph containing a source node, sink node, a node for each student, a node for each course, and a node for each course TA slot. The assignments producing the highest total weight is written to a csv file. For each student, the output contains
- Course the student is assigned to, or "unassigned"
- Total weight of the edge
- Ranking of the match indicated in the student's and professor's responses
- Rank score: Z-score of the student's ranking among all of the student's rankings for qualified courses, and similar for professor
- Match score: Z-score of the edge weight among all possible edges for the student, and similar from the course perspective

### Additional TA
Assuming an extra TA is aquired for a particular course, so one of the course's slots is filled but unavailable, calculates the amount by which the total weight in the new best matching changes.

### Removing TA
Assuming a student is removed from consideration from the matching, calculates the amount by which the total weight in the new best matching decreases.

## Example Usage

In the project directory root, running the following will perform the algorithm on the test inputs and save the outputs in `./test/outputs`:
```
python matching.py --path test/
```
