from typing import List, Tuple
import pandas as pd
import json
import csv
import os


class Mentor:
    def __init__(self, row):
        self.full_name = getattr(row, 'What_is_your_full_name', None)
        self.gender = getattr(row, 'Are_you_a_brother_or_a_sister', None)
        self.short_code = getattr(row, 'Please_enter_your_Imperial_shortcode', None)
        self.phone_number = getattr(row, 'Please_enter_your_phone_number', None)
        self.course = getattr(row, 'What_course_are_you_studying', None)
        # Also store lowercased version for easier matching
        self.course_lower = self.course.lower() if self.course else ""
        self.year = getattr(row, 'What_year_of_study_are_you_in', None)
        self.stem_muslims_mentor_before = getattr(row, 'Have_you_been_a_STEM_Muslims_Mentor_before', None)
        self.max_students = 6 if getattr(row, 'How_many_students_are_you_able_to_mentor', None) in ['5+', '6'] else int(getattr(row, 'How_many_students_are_you_able_to_mentor', 0))
        self.current_students = getattr(row, 'current_student_numbers', 0)
            
    def __lt__(self, other):
        return (
            self.current_students < other.current_students
            and self.current_students < self.max_students
        )


class Mentee:
    def __init__(self, row) -> None:
        self.full_name = getattr(row, 'What_is_your_full_name', None)
        self.gender = getattr(row, 'Are_you_a_brother_or_a_sister', None)
        self.email = getattr(row, 'Please_enter_your_email', None)
        self.phone_number = getattr(row, 'Please_enter_your_phone_number', None)
        self.current_study_place = getattr(row, 'Where_are_you_currently_studying', None)
        self.a_levels = getattr(row, 'What_Alevels_are_you_currently_studying', None)
        import re

        try:
            raw_subjects = getattr(
                row, 
                'What_subjects_are_you_interested_in_studying_at_university', 
                None
            )

            if not raw_subjects:  # if first attribute is None/empty, try the other
                raw_subjects = getattr(
                    row, 
                    'What_subjects_/_courses_are_you_interested_in_studying_at_university', 
                    None
                )
            if "Business/EFDS (Economics, Finance and Data Science)" in raw_subjects:
                raw_subjects = raw_subjects.replace("Business/EFDS (Economics, Finance and Data Science)", "Business/EFDS")
            if raw_subjects:
                self.interested_subjects = [
                    item.strip() 
                    for item in re.split(r"[;]", raw_subjects) 
                    if item.strip()
                ]
            else:
                self.interested_subjects = None

        except Exception as e:
            print(f"Error parsing interested_subjects for mentee {self.full_name}: {e}")
            self.interested_subjects = None

        self.why_interested = getattr(row, 'Why_are_you_interested_in_studying_this_subject_/_field', None)
        self.areas_of_advice = getattr(row, 'What_areas_of_the_UCAS_process_are_you_looking_for_support_with', None)
        self.considering_imperial = getattr(row, 'Are_you_considering_applying_to_Imperial', None)


class MentorList:
    def __init__(self, csv_path) -> None:
        self.exception_students: List[Mentee] = []

        df = pd.read_csv(csv_path)
        self.mentor_pairings: List[Tuple[Mentor, List[Mentee]]] = [
            (Mentor(row), []) for row in df.itertuples()
        ]

    def pair_mentee(self, mentee: Mentee):

        for mentor, mentors_mentees in self.mentor_pairings:
            try:
                matched_subject = None
                for subj in mentee.interested_subjects:
                    if subj and subj.lower() in (mentor.course_lower if mentor.course_lower else ""):
                        matched_subject = subj
                        break
                if (
                    matched_subject is not None
                    and mentor.gender == mentee.gender
                    and mentor.max_students > mentor.current_students
                ):
                    print(
                        f"{mentee.full_name} paired with {mentor.full_name} Mentor does {mentor.course}, mentee interested in {mentee.interested_subjects}\n----------------------------------------------------------------------------"
                    )
                    mentee.matched_course = matched_subject
                    mentors_mentees.append(mentee)
                    self.mentor_pairings.sort()
                    mentor.current_students += 1
                    return
            except Exception as e:
                print(f"Error in pairing {mentee.full_name} with {mentor.full_name}: {e}")
                break
        if mentee.interested_subjects != [""]:
            print(f"{mentee.full_name} could not be paired, \ngender: {mentee.gender}\ncourses: {mentee.interested_subjects}\n----------------------------------------")
            self.exception_students.append(mentee)

    def make_mentor_mentee_json(self):
        json_mentor_mentee_pairs_data = []
        json_unpaired_mentor_data = []
        json_unpaired_mentee_data = [
            {"mentee": vars(mentee)} for mentee in self.exception_students
        ]
        for (
            mentor,
            mentees,
        ) in self.mentor_pairings:
            if mentees:
                entry = {
                    "mentor": vars(mentor),
                    "mentees": [vars(mentee) for mentee in mentees],
                }
                json_mentor_mentee_pairs_data.append(entry)
            else:
                entry = {"mentor": vars(mentor)}
                json_unpaired_mentor_data.append(entry)

        json_mentor_mentee_pairs_string = json.dumps(
            json_mentor_mentee_pairs_data, indent=4
        )
        json_unpaired_mentee_string = json.dumps(json_unpaired_mentee_data, indent=4)
        json_unpaired_mentor_string = json.dumps(json_unpaired_mentor_data, indent=4)

        with open("mentor_mentee_pairings.json", "w") as f:
            f.write(json_mentor_mentee_pairs_string)
        with open("unpaired_mentors.json", "w") as f:
            f.write(json_unpaired_mentor_string)
        with open("unpaired_mentees.json", "w") as f:
            f.write(json_unpaired_mentee_string)

        unpaired_mentors = [] 
        for (mentor, mentees) in self.mentor_pairings:
            if not mentees:
                unpaired_mentors.append(list(vars(mentor).values())[:-1])
        
        with open(os.path.join('csvs', 'mentors.csv'), 'r') as f:
            reader = csv.reader(f)
            first_row = next(reader)

        with open(os.path.join('csvs', 'mentors.csv'), 'w') as f:
            writer = csv.writer(f)
            writer.writerow(first_row)
            for (mentor, _) in self.mentor_pairings:
                writer.writerow(list(vars(mentor).values()))

        #Gets the top row of the csv
        with open(os.path.join('csvs', 'mentees.csv'), 'r') as f:
            reader = csv.reader(f)
            first_row = next(reader)
   
        mentees_df = pd.read_csv(os.path.join('csvs', 'mentees.csv')).rename(columns={'What_is_your_full_name' : 'full_name'})
        mentees_df['full_name'] = mentees_df['full_name'].str.strip()
        
        unpaired_mentees_df = pd.DataFrame([vars(mentee) for mentee in self.exception_students])
        unpaired_mentees_df['full_name'] = unpaired_mentees_df['full_name'].str.strip()
        merged_df = pd.merge(mentees_df, unpaired_mentees_df[['full_name']], on='full_name', how='inner')

        merged_df.to_csv(os.path.join('csvs', 'unpaired_mentees.csv'), index=False)




