import json
import os
import shutil
import re

full_name = lambda ment__: ment__['full_name'].strip().title()
first_name = lambda ment__: ment__['full_name'].split()[0].strip().capitalize()

def mentor_email(mentor):
    if re.match(r'^[a-zA-Z]{2,}[0-9]{2,}$', mentor['short_code']):
        return f"{mentor['short_code']}@ic.ac.uk"   # fixed: missing '@'
    elif re.match(r'.*@(ic\.ac\.uk|imperial\.ac\.uk)$', mentor['short_code']):
        return mentor['short_code']
    else:
        return "REPLACE WITH CORRECT EMAIL"

def clean_email(email):
    """Ensure no duplicate domains like @gmail.com@gmail.com"""
    email = email.strip()
    if email.count("@") > 1:  # if user accidentally entered twice
        first, *rest = email.split("@")
        return first + "@" + rest[-1]   # keep only last part
    return email

email_format = lambda mentee, mentor, his_her='their': (
f"""Assalamu Alaykum,

Firstly, thank you {full_name(mentee)}, for signing up for the UCAS Support Service. We are committed to assisting you throughout this pivotal phase of your academic journey. We hope the support you receive over the next coming weeks will be fruitful and beneficial inshaAllah. 

{first_name(mentee)}, I'm pleased to introduce you to {full_name(mentor)}, who is a {mentor['year']} Year studying {mentor['course']} at Imperial College London who will be your dedicated mentor, guiding you throughout your UCAS Application process. 

{first_name(mentor)}, I'd like to introduce to you {full_name(mentee)}, who is a diligent sixth form student also aspiring to study {mentor['course']} and we found you are a perfect match to support {first_name(mentee)} in {his_her} endeavours. 

I hope you both get acquainted with each other and make this journey a smooth and rewarding experience.
If you require any support or anything else, please feel free to get in touch.

Mentor's email: {mentor_email(mentor)}

Mentee's email: {clean_email(mentee['email'])}

Wishing you all the best,
STEM Muslims"""
)

def generate_emails():
    if os.path.isdir("emails"):
        shutil.rmtree("emails")
    os.mkdir("emails")

    with open("mentor_mentee_pairings.json", "r", encoding="utf-8") as f:
        pairings = json.load(f)
    
    total_emails = 0
    for pair in pairings:
        mentor = pair["mentor"]
        mentees = pair["mentees"]
        
        folder = os.path.join(os.getcwd(), 'emails', full_name(mentor).replace(' ', '_'))
        os.makedirs(folder, exist_ok=True)
        os.mkdir(os.path.join(folder, "mentees_details"))

        # Save mentor details
        with open(os.path.join(folder, f"mentor_details.json"), 'w', encoding="utf-8") as f:
            json.dump(mentor, f, indent=4, ensure_ascii=False)

        he_she = "he" if mentor["gender"] == "Brother" else "she"
        his_her = "his" if mentor["gender"] == "Brother" else "her"

        for mentee in mentees:
            # Save the actual email text
            email_file = os.path.join(
                folder,
                f"{full_name(mentor).replace(' ', '_')}_{full_name(mentee).replace(' ', '_')}.txt"
            )
            with open(email_file, 'w', encoding="utf-8") as f:
                f.write(email_format(mentee, mentor, his_her))
            
            # Save mentee details separately
            mentee_file = os.path.join(folder, 'mentees_details', f"{full_name(mentee).replace(' ', '_')}.json")
            with open(mentee_file, 'w', encoding="utf-8") as f:
                json.dump(mentee, f, indent=4, ensure_ascii=False)

            total_emails += 1

    print(f" Generated {total_emails} email drafts.")

if __name__ == "__main__":
    generate_emails()
