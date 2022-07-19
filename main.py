import pandas

courses_list = {
    1: "Matematik",
    2: "Fizik",
    3: "Kimya",
    4: "Lineer Cebir"
}

note_range = {
    "A": "100-90",
    "B": "89-80",
    "C": "80-70",
    "D": "70-50",
    "F": "50-0"
}
select_process_list = {
    1: "Add to Student",
    2: "Status to Calculate",
    3: "Export to EXCEL",
    4: "Quit"
}
student_list = {}
course_id, process_id = 0, 0


def select_course():
    for key, value in courses_list.items():
        print(f"{key} for {value}", end="\n")
    global course_id
    course_id = int(input("Please, Enter one Course : "))
    while True:
        if not course_id in courses_list.keys():
            print("*************************************** ", end="\n")
            print("Please use defined code ", end="\n")
            print("*************************************** ", end="\n")
            course_id = int(input("Please, Enter one Course : "))
        else:
            break
    print("***************************************", end="\n")
    select_process()


def select_process():
    for key, value in select_process_list.items():
        print(f"{key} for {value}", end="\n")
    global process_id
    process_id = int(input("Please, Enter one Process : "))
    while True:
        if not process_id in select_process_list.keys():
            print("*************************************** ", end="\n")
            print("Please use defined code ", end="\n")
            print("*************************************** ", end="\n")
            process_id = int(input("Please, Enter one Process : "))
        else:
            break
    if select_process_list[process_id] == "Add to Student":
        add_to_student()
    elif select_process_list[process_id] == "Status to Calculate":
        calculate_status()
    elif select_process_list[process_id] == "Export to EXCEL":
        write_to_excel_datasheet()
    else:
        print("Terminated of Process")


def add_to_student():
    process = True
    while process:
        name = input("Please, enter to name :")
        lastname = input("Please, enter to lastname :")
        student_number = input("Please, enter to student_number(22xxxx) :")
        exam_score = input("Please, enter to exam score :")
        if student_number in student_list:
            print(f"Student number of {student_number} is registered.")
        else:
            student_list[student_number] = {
                "name": name,
                "lastname": lastname,
                "course": courses_list[course_id],
                "exam_score": int(exam_score),
                "letter_grade": "",
                "status": ""
            }
        if input("Continue for \"C\" , Quit for \"Q\" : ").upper() == "C":
            process = True
        else:
            process = False
    print("Registration Successful.")
    select_process()


def calculate_status():
    get_lenght_student()
    for student_number in student_list.keys():
        if student_list[student_number]["exam_score"] >= 90 and student_list[student_number]["exam_score"] <= 100:
            student_list[student_number]["letter_grade"] = "A"
            student_list[student_number]["status"] = "PASSED"
        elif student_list[student_number]["exam_score"] >= 80 and student_list[student_number]["exam_score"] < 90:
            student_list[student_number]["letter_grade"] = "B"
            student_list[student_number]["status"] = "PASSED"
        elif student_list[student_number]["exam_score"] >= 70 and student_list[student_number]["exam_score"] < 80:
            student_list[student_number]["letter_grade"] = "C"
            student_list[student_number]["status"] = "PASSED"
        elif student_list[student_number]["exam_score"] >= 50 and student_list[student_number]["exam_score"] < 70:
            student_list[student_number]["letter_grade"] = "D"
            student_list[student_number]["status"] = "PASSED"
        else:
            student_list[student_number]["letter_grade"] = "F"
            student_list[student_number]["status"] = "STAYED"
    select_process()


def write_to_excel_datasheet():
    get_lenght_student()
    information_list, student_number_list = [], []
    for student_number in student_list:
        student_number_list.append(student_number)
        information_list.append([student_list[student_number]["name"], student_list[student_number]["lastname"],
                                 student_list[student_number]["course"], student_list[student_number]["exam_score"],
                                 student_list[student_number]["letter_grade"], student_list[student_number]["status"]])

    dataframe = pandas.DataFrame(information_list,
                                 index=student_number_list,
                                 columns=["Name", "Lastname", "Course", "Exam Score", "Letter Grade", "Status"])
    with pandas.ExcelWriter("WorkShop Sheet.xlsx") as writer:
        dataframe.to_excel(writer, sheet_name="Calculate of Course")


def get_lenght_student():
    if not len(student_list) > 0:
        print("*************************************** ", end="\n")
        print("Please before add student data")
        print("*************************************** ", end="\n")
        select_process()


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    select_course()
