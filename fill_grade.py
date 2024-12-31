# This code is designed to fille the score sheet in NTU Cool
# You can choose either "normal" or "revise" mode for data input
# (mode can be set at the start of entering data)


import pandas as pd
import re
import numpy as np



name = "./test.xlsx"
score_title = "期末考"
mode = "calculate"     # "fill_in",  "calculate"

df = pd.read_excel(name)


# Define a regular expression pattern to match English characters
pattern = re.compile(r'[a-zA-Z(),\s-]')


# define
def read_grade(mode):
    if mode == "fill_in":
        grade = input(">>>  Grade  : ")
        if grade.isdigit():
            grade = int(grade)
            return grade
        else:
            return "exit"
    elif mode == "calculate":
        flag = 1
        while (flag == 1):
            print()
            input_data = input("Enter a grade array: ")
            # stop the while loop
            if input_data == "end":
                flag = 0
                break

            input_data = input_data.replace("  ", " ").replace("   ", " ")
            if input_data[0] == " ":
                input_data = input_data[1:]
            if input_data[-1] == " ":
                input_data = input_data[:len(input_data) - 1]
                # print(input_data)
            input_data = input_data.replace("  ", " ").replace("   ", " ")

            grade_list = input_data.split(" ")
            
            grade_list = [int(x) for x in grade_list]

            return np.sum(grade_list)


revise_mode = False
flag = 1
while (flag == 1):

    print()
    input_data = input("Enter a string: ")
    # stop the while loop
    if input_data == "end" or input_data == "exit":
        flag = 0
        break
    elif input_data == "revise":
        print("<<<Change to \"revise\" mode.>>>")
        input_data = input("Enter a string again: ")
        if input_data  == "":
            continue
        revise_mode = True
    elif input_data == "normal":
        print("<<<Change to \"normal\" mode.>>>")
        input_data = input("Enter a string again: ")
        if input_data  == "":
            continue
        revise_mode = False
    elif input_data == "":
        continue
    # elif input_data == "noscore" or input_data == "nograde":
    #     np.where(isinstance(df[score_title].values, str))
    #     print(a)
    #     student_no_grade = df["Student ID"][np.isnan(list(df[score_title].values))]
    #     print(student_no_grade)
    #     continue
    
    # select data
    index_bool = df["Student ID"].str.find(input_data) > -1
    student_candidate = df["Student ID"][index_bool]
    if revise_mode:
        grade_select = np.invert(np.isnan(list(df[score_title][index_bool])))
    else:
        grade_select = np.isnan(list(df[score_title][index_bool]))
    student_candidate = student_candidate[grade_select]

    # check student
    if len(student_candidate) > 1:
        print()
        print(f"There are \"{len(student_candidate)}\" possible students: ")
        # print(df["Name"][df["Student ID"].str.find(input_data) > 0])
        
        name_list = df["Name"][df["Student ID"].str.find(input_data) > 0][grade_select]
        candidate_name = [re.sub(pattern, '', string) for string in name_list.values]
        for i in range(len(candidate_name)):
            print(f"{i+1}: {candidate_name[i]}")

        # choose again
        choose_student = input("Which one is the corret student? ")   # Type number of index "in normal order" rather then in CS (starting at 1)
                                                                      # If there is no correct student, type "0"
        if choose_student == "":
            continue
        else:
            choose_student = int(choose_student) - 1

        if choose_student < 0 or choose_student > len(student_candidate) -1:
            print("Out of index")
            continue
        student_index = student_candidate.index[int(choose_student)]
    elif len(student_candidate) == 1:
        student_index = student_candidate.index[0]
    else:
        print("No student match!")
        continue
    
    print(f"The index is {student_index}")
    print(f">>>  {df['Name'][student_index][:4]}: {df['Student ID'][student_index]}")
    if revise_mode:
        print(f"     (old grade: {df[score_title][student_index]})")

    # input grade
    grade = read_grade(mode = mode)
    if mode == "calculate":
        print(f"Fill score {grade}")

    if grade == "exit":
        continue



    # Write data
    # warnings.simplefilter(action="ignore", category=SettingWithCopyWarning)
    with pd.option_context('mode.chained_assignment', None):
        df[score_title][student_index] = grade
        df.to_excel(name, index=False)


    

print()
print("End of filling data")
