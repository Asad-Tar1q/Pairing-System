import json
import os
import shutil
import re

# ---------- helpers ----------

def norm_name(s: str) -> str:
    """Title-case a name and collapse to safe filename (A-Z, a-z, 0-9, _)."""
    if not s:
        return ""
    titled = str(s).strip().title()
    return re.sub(r"[^A-Za-z0-9_]+", "_", titled)

def full_name(person: dict) -> str:
    return str(person.get("full_name", "")).strip().title()

def first_name(person: dict) -> str:
    return str(person.get("full_name", "").split()[0]).strip().capitalize()

def mentor_email(mentor: dict) -> str:
    sc = str(mentor.get("short_code", ""))
    # e.g. ab123@ic.ac.uk or ab12345
    if re.fullmatch(r"[A-Za-z]{2,}\d{2,}", sc):
        return f"{sc}@ic.ac.uk"
    # Proper domain check for ic/imperial
    if re.fullmatch(r".*@(?:ic|imperial)\.ac\.uk", sc):
        return sc
    return "REPLACE WITH CORRECT EMAIL"

def email_format(mentee, mentor, his_her="their") -> str:
    return f"""Assalamu Alaykum,  

We are pleased to introduce you to one another as part of the UCAS Mentorship Scheme. This programme is designed to provide guidance and encouragement during a key stage of the university application journey, and we pray that the experience will be beneficial for you both, inshaAllah.  

Here's a little about each of you:  
- {full_name(mentee)} is a sixth form student preparing for university applications.  
- {full_name(mentor)} is a {mentor['year']} Year {mentor['course']} student at Imperial College London and will be supporting {first_name(mentee)} throughout the UCAS process.  

We encourage you to get in touch and begin working together, making the most of this opportunity to share knowledge and build a supportive mentoring relationship. If you require any assistance or have any questions, please don't hesitate to contact us.  

Mentor's email: {mentor_email(mentor)}  
Mentee's email: {mentee['email']}  

Wishing you both every success,  
STEM Muslims
"""
# Assalamu Alaykum,

# Firstly, thank you {full_name(mentee)}, for signing up for the UCAS Support Service. We are committed to assisting you throughout this pivotal phase of your academic journey. We hope the support you receive over the next coming weeks will be fruitful and beneficial inshaAllah. 

# {first_name(mentee)}, I'm pleased to introduce you to {full_name(mentor)}, who is a {mentor['year']} Year studying {mentor['course']} at Imperial College London who will be your dedicated mentor, guiding you throughout your UCAS Application process. 

# {first_name(mentor)}, I'd like to introduce to you {full_name(mentee)}, who is a diligent sixth form student also aspiring to study {mentor['course']} and we found you are a perfect match to support {first_name(mentee)} in {his_her} endeavours.

# I hope you both get acquainted with each other and make this journey a smooth and rewarding experience.
# If you require any support with setting up meetings or anything else, please feel free to get in touch. 

# Mentor's email: {mentor_email(mentor)}
# Mentee's email: {mentee['email']}

# Wishing you all the best,
# STEM Muslims
# """

# ---------- main ----------

def generate_emails():
    import pandas as pd

    # Start fresh each run
    if os.path.isdir("emails"):
        shutil.rmtree("emails")
    os.mkdir("emails")

    # Read pairings
    df = pd.read_excel("pairings_export.xlsx", sheet_name="Pairings")

    # Group by mentor attributes
    grouped = df.groupby(
        ["Mentor Name", "Mentor Short Code", "Mentor Year", "Mentor Course", "Mentor Gender"],
        dropna=False,
    )

    total_emails = 0

    for (mentor_name, mentor_shortcode, mentor_year, mentor_course, mentor_gender), group in grouped:
        mentor = {
            "full_name": str(mentor_name),
            "short_code": str(mentor_shortcode),
            "year": str(mentor_year),
            "course": str(mentor_course),
            "gender": (
                "Brother" if "brother" in str(mentor_gender).lower()
                else "Sister" if "sister" in str(mentor_gender).lower()
                else "N/A"
            ),
        }

        # Folder per mentor
        mentor_folder_name = norm_name(full_name(mentor))
        folder = os.path.join(os.getcwd(), "emails", mentor_folder_name)
        os.makedirs(folder, exist_ok=True)

        # Create mentees_details, but don't error if it exists
        mentees_dir = os.path.join(folder, "mentees_details")
        os.makedirs(mentees_dir, exist_ok=True)

        # Persist mentor details
        with open(os.path.join(folder, "mentor_details.json"), "w", encoding="utf-8") as f:
            json.dump(mentor, f, indent=4, ensure_ascii=False)

        # Pronouns (kept for future-proofing if you use them)
        he_she = "he" if mentor["gender"] == "Brother" else "she"
        his_her = "his" if mentor["gender"] == "Brother" else "her"
        him_her = "him" if mentor["gender"] == "Brother" else "her"

        # Create emails + mentee jsons
        for _, mentee_row in group.iterrows():
            mentee = {
                "full_name": str(mentee_row["Mentee Name"]),
                "email": str(mentee_row["Mentee Email"]),
            }

            # Email text file
            email_filename = f"{norm_name(full_name(mentor))}_{norm_name(full_name(mentee))}.txt"
            mentee_email_path = os.path.join(folder, email_filename)
            with open(mentee_email_path, "w", encoding="utf-8") as f:
                f.write(email_format(mentee, mentor, his_her))

            # Save mentee details
            mentee_json = os.path.join(mentees_dir, f"{norm_name(full_name(mentee))}.json")
            with open(mentee_json, "w", encoding="utf-8") as f:
                json.dump(mentee, f, indent=4, ensure_ascii=False)

            total_emails += 1

    print(total_emails)

if __name__ == "__main__":
    generate_emails()
