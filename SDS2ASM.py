"""\
This script converts Microsoft School Data Sync Classic zip files to Apple School Manager zip files

Usage: sds2asm.py input.zip [--quiet]
--quiet/-q : do not prompt to overwrite output files
--help/-h  : display this message
"""


import argparse
import zipfile
import io
import csv
import os
import sys
import datetime

# Function to read Microsoft SDS formatted files from the input zip
def read_sds_files(input_zip):
    expected_files = [
        'School.csv',
        'Section.csv',
        'Student.csv',
        'Teacher.csv',
        'TeacherRoster.csv',
        'StudentEnrollment.csv'
    ]
    
    SDSdata = {}
    
    print("Begin processing",input_zip)
    
    with zipfile.ZipFile(input_zip, 'r') as zip_ref:
        file_list = zip_ref.namelist()
        
        # Filter out macOS-specific hidden files
        file_list = [file_name for file_name in file_list if '__MACOSX' not in file_name]
        
        
        # Check if the number of CSV files is as expected
        if len(file_list) != len(expected_files):
            print("Error: The input zip file does not contain the expected number of CSV files.")
            print(file_list)
            sys.exit(1)
            
        for expected_file in expected_files:
            if expected_file not in file_list:
                print(f"Error: The input zip file is missing the '{expected_file}' CSV file.")
                sys.exit(1)
                
            file_name_lower = expected_file.lower()
            with zip_ref.open(expected_file, 'r') as file:
                file_content = file.read().decode('utf-8')
                
                # Check if the CSV file is empty
                if len(file_content.strip()) == 0:
                    print(f"Error: The '{expected_file}' CSV file is empty.")
                    sys.exit(1)
                    
                # Check the encoding of the CSV file
                if '\ufeff' in file_content:  # Check for UTF-8 BOM
                    print(f"Warning: The '{expected_file}' CSV file contains UTF-8 BOM. Consider saving it without BOM.")
                    file_content = file_content.replace('\ufeff', '')
                
                SDSdata[file_name_lower.replace('.csv', '')] = list(csv.DictReader(io.StringIO(file_content)))
                
    return SDSdata

# Function to generate the locations.csv data
def generateLocations(sds_locations, asm_data):
    if not sds_locations:
        print("Error: No data for locations found.")
        sys.exit(1)

    asm_locations = []
    # DEBUG print("Total number of records in locations_data:", len(locations_data))


    for location_entry in sds_locations:
        asm_location = {
            'location_id': location_entry.get('SIS ID', ''),
            'location_name': location_entry.get('Name', '')
        }
        asm_locations.append(asm_location)
        # DEBUG print("Generated location entry:", asm_location)
    
    if not asm_locations:
        print("Error: No locations found in the source data. Where is your school?")
        sys.exit(1)
        
    # Add the generated data to ASMdata
    asm_data['locations'] = asm_locations

# Function to generate the students.csv data
def generateStudents(sds_students, asm_data):
    asm_students = []

    for sds_student in sds_students:
        email_address = sds_student['Username']
        student_number = sds_student['Student Number']

        # Calculate grade level based on email address
        class_year = int(email_address.split('@')[0][-2:])
        current_year = datetime.datetime.now().year % 100
        if datetime.datetime.now().month >= 8:
            current_year += 1
        grade_level = 12 + (current_year - class_year)

        # Determine password_policy based on grade level
        password_policy = 8 if grade_level >= 3 else 4

        asm_student = {
            'person_id': sds_student['SIS ID'],
            'person_number': student_number,
            'first_name': sds_student['First Name'],
            'middle_name': '',
            'last_name': sds_student['Last Name'],
            'grade_level': grade_level,
            'email_address': email_address,
            'sis_username': email_address,
            'password_policy': password_policy,
            'location_id': sds_student['School SIS ID']
        }
        asm_students.append(asm_student)

    if not asm_students:
        print("Error: No students found in the source data. That's an empty school.")
        sys.exit(1)
        
    asm_data['students'] = asm_students

# Function to generate the staff.csv data
def generateStaff(sds_staff, asm_data):
    asm_staff = []

    for sds_staff_member in sds_staff:
        asm_staff_member = {
            'person_id': sds_staff_member['SIS ID'],
            'person_number': sds_staff_member['SIS ID'],
            'first_name': sds_staff_member['First Name'],
            'middle_name': '',
            'last_name': sds_staff_member['Last Name'],
            'email_address': sds_staff_member['Username'],
            'sis_username': sds_staff_member['Username'],
            'location_id': sds_staff_member['School SIS ID']
        }
        asm_staff.append(asm_staff_member)

    if not asm_staff:
        print("Error: No staff found in the source data. Who will teach the children?")
        sys.exit(1)
        
    asm_data['staff'] = asm_staff
    
def generateCourses(sds_courses, asm_data):
    asm_courses = []

    unique_course_ids = set()  # To track unique course IDs

    for sds_course in sds_courses:
        course_id = sds_course['Course SIS ID']
        
        # Check if the course ID is already added, skip if it's a duplicate
        if course_id in unique_course_ids:
            continue
        
        asm_course = {
            'course_id': course_id,
            'course_number': course_id,
            'course_name': sds_course['Course Name'],
            'location_id': sds_course['School SIS ID']
        }
        asm_courses.append(asm_course)
        
        # Add the course ID to the set of unique course IDs
        unique_course_ids.add(course_id)

    if not asm_courses:
        print("Error: No courses found in the source data. What are you teaching them?")
        sys.exit(1)
        
    asm_data['courses'] = asm_courses

def generateClasses(sds_sections, sds_teacherroster, asm_data):
    asm_classes = []
    
    # Create a dictionary to map section IDs to instructor IDs
    instructor_mapping = {row['Section SIS ID']: row['SIS ID'] for row in sds_teacherroster}
    
    for sds_section in sds_sections:
        section_sis_id = sds_section['SIS ID']
        instructor_id = instructor_mapping.get(section_sis_id, '')  # Get instructor ID from the mapping, default to '' if not found
        
        asm_class = {
            'class_id': sds_section['SIS ID'],
            'class_number': sds_section['Section Name'],
            'course_id': sds_section['Course SIS ID'],
            'instructor_id': instructor_id,
            'location_id': sds_section['School SIS ID']
        }
        asm_classes.append(asm_class)
    
    if not asm_classes:
        print("Error: No classes found in the source data. When are you teaching?")
        sys.exit(1)
        
    asm_data['classes'] = asm_classes
    

def generateRosters(sds_enrollments, ASMdata):
    asm_rosters = []
    
    for sds_enrollment in sds_enrollments:
        asm_roster = {
            'roster_id': f"{sds_enrollment['SIS ID']}.{sds_enrollment['Section SIS ID']}",
            'class_id': sds_enrollment['Section SIS ID'],
            'student_id': sds_enrollment['SIS ID']
        }
        asm_rosters.append(asm_roster)
        
    if not asm_rosters:
        print("Error: No schedules found in the source data. Look who has a very light load!")
        sys.exit(1)
        
    ASMdata['rosters'] = asm_rosters

def output_asm_files(output_zip, ASMdata, quiet_mode=False):
    # Check if the output file already exists
    if os.path.exists(output_zip):
        if quiet_mode:
            # Overwrite without prompting if in quiet mode
            print("Quiet Mode Enabled: Overwriting existing output file", output_zip)
        else:
            # Prompt the user for permission to delete the existing file
            response = input("Output file already exists. Do you want to overwrite it? (y/n): ")
            if response.lower() != 'y':
                print("Exiting without overwriting.")
                sys.exit(1)

        # Delete the existing output file
        os.remove(output_zip)
        
    with zipfile.ZipFile(output_zip, 'a') as zip_ref:
        for table_name, table_data in ASMdata.items():
            csv_filename = f'{table_name}.csv'
            print("Writing CSV:", csv_filename, "(",len(table_data),")")  # Print the CSV filename
            
            with io.StringIO() as csv_buffer:
                csv_writer = csv.DictWriter(csv_buffer, fieldnames=table_data[0].keys())
                csv_writer.writeheader()  # Write headers only once
                
                rows = [row for row in table_data]
                csv_writer.writerows(rows)  # Write all rows at once

                # Convert CSV data to bytes
                csv_bytes = csv_buffer.getvalue().encode('utf-8')
                
                # Write CSV data to the zip file
                zip_ref.writestr(csv_filename, csv_bytes)

# Main function to run the script
def main(input_zip, quiet_mode):
    sds_data = read_sds_files(input_zip)

    ASMdata = {}  # Initialize ASMdata

    # Generate each ASM CSV file
    generateLocations(sds_data.get('school'), ASMdata)
    generateStudents(sds_data.get('student'), ASMdata)
    generateStaff(sds_data.get('teacher'), ASMdata)
    generateCourses(sds_data.get('section'), ASMdata)
    generateClasses(sds_data.get('section'), sds_data.get('teacherroster'), ASMdata)
    generateRosters(sds_data.get('studentenrollment'), ASMdata)
    
    # Output all ASM CSVs
    output_asm_files(input_zip.replace('.zip', '_asm.zip'), ASMdata, quiet_mode)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Convert Microsoft SDS v1 format to Apple School Manager zip format (ASM2SDS)")
    parser.add_argument("input_zip", metavar="input.zip", type=str, help="Input SDS zip file")
    parser.add_argument("-q", "--quiet", action="store_true", help="Suppress prompts and overwrite existing output file automatically")

    args = parser.parse_args()

    if args.input_zip is None:
        print("Error: Missing input file.")
        parser.print_help()
        sys.exit(1)

    main(args.input_zip, args.quiet)
    