# weight added to favorite-favorite matches
FAVORITE_FAVORITE = 8
# weight added to matches where student has previously TAed course
PREVIOUS = 2
# weight added to matches where professor is the student's advisor
ADVISORS = 2
# weight added to good (student) - favorite (instructor) matches
STUDENT_GOOD_INSTRUCTOR_FAVORITE = 7
# weight added to favorite (student) - neutral (instructor) matches
STUDENT_FAVORITE_INSTRUCTOR_NEUTRAL = 2
# weight to fill first course slot
DEFAULT_FIRST_FILL = 20
# weight to fill any course slot
DEFAULT_BASE_FILL = 5
# weight for giving a student a TA position
DEFAULT_ASSIGN = 5
# default value for JOIN if no value is supply
DEFAULT_JOIN = 3.0
# multiplier of JOIN value to add to weight
JOIN_MULTIPLIER = 12.0
# default value for BANK if no value is supply
DEFAULT_BANK = 3.0
# multiplier of BANK value to add to weight
BANK_MULTIPLIER = -10.0
# value by which to increase all MSE students
MSE_BOOST = 20.0
# weight added to (student, course) pairs for all courses in a student's OK list
OKAY_COURSE_PENALTY = -6.0
# weight added to (student, course) pairs that were in the specified previous matching execution
PREVIOUS_MATCHING_BOOST = 0.05
# weight added to a student per course they ranked as 'Favorite' or 'Good'
BOOST_PER_COURSE_STUDENT_RANKED = 0.05
# weight added to a course per student they ranked as 'Favorite'
BOOST_PER_FAVORITE_STUDENT = 0.15
