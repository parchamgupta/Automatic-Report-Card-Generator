#importing necessary libraries
import pandas as pd
import os
from fpdf import FPDF
import matplotlib.pyplot as plt

#class declaration
class student:
    #variables
    df = pd.DataFrame()

    #constructor
    def __init__(self, datafile):
        #reading the excel file into a pandas dataframe
        self.df = pd.read_excel(datafile, skiprows=1)

    #data preprocessing
    def preprocess(self):
        #filling NULL values
        self.df['What you marked'] = self.df['What you marked'].fillna('Unattempted')
        #reading the images
        photos_ls = os.listdir("./Student Photos/")
        #calculating the number of students and the number of questions
        no_students = len(self.df['Registration Number'].unique())
        no_ques = int(self.df.shape[0]/no_students)
        #returning to the main function
        return no_students, no_ques, photos_ls

    #getting student details
    def student_details(self, counter, photo_ls, num):
        #counter for student count
        row = counter * num
        #extracting student personal details from excel file
        round = self.df.iloc[row]['Round'].astype('str')
        reg_no = self.df.iloc[row]['Registration Number'].astype('str')
        f_name = self.df.iloc[row]['First Name ']
        l_name = self.df.iloc[row]['Last Name ']
        grade = self.df.iloc[row]['Grade '].astype('str')
        gender = self.df.iloc[row]['Gender']
        dob = self.df.iloc[row]['Date of Birth ']
        school = self.df.iloc[row]['Name of School ']
        city_residence = self.df.iloc[row]['City of Residence']
        country_residence =  self.df.iloc[row]['Country of Residence']
        date = self.df.iloc[row]['Date and time of test']
        photo = [ph for ph in photo_ls if (f_name + " " + l_name + ".jpg") == ph]
        comment = self.df.iloc[row]['Qualification']
        #returning to the main function
        return photo[0], reg_no, f_name, l_name, dob, gender, grade, school, city_residence, country_residence, date, comment, round

    #getting student marks
    def student_marks(self, counter):
        #extracting student marks and scorecard related data from excel file
        ques_no = self.df.iloc[counter]['Question No.']
        student_answer = self.df.iloc[counter]['What you marked']
        correct_answer = self.df.iloc[counter]['Correct Answer']
        student_marks = self.df.iloc[counter]['Your score'].astype('str')
        max_marks = self.df.iloc[counter]['Score if correct'].astype('str')
        outcome = self.df.iloc[counter]['Outcome (Correct/Incorrect/Not Attempted)']
        #returning to the main function
        return ques_no, student_answer, correct_answer, student_marks, max_marks, outcome

    #getting comparative data of the student with the rest of the world
    def esr_report(self, counter):
        #extracting data from excel file
        data1 = self.df.iloc[counter]['% of students\nacross the world\nwho attempted\nthis question'].astype('str')
        data2 = self.df.iloc[counter]['% of students (from\nthose who attempted\nthis ) who got it\ncorrect'].astype('str')
        data3 = self.df.iloc[counter]['% of students\n(from those who\nattempted this)\nwho got it\nincorrect'].astype('str')
        data4 = self.df.iloc[counter]['World Average\nin this question\n'].astype('str')
        #returning to main function
        return  data1, data2, data3, data4

    #getting data for plotting graphs
    def overview(self, counter, num):
        #counter for student count
        row = counter * num
        #extracting data from excel file
        avg = self.df.iloc[row]['Average score of all students across the World'].astype('str')
        median = self.df.iloc[row]['Median score of all students across the World'].astype('str')
        mode = self.df.iloc[row]['Mode score of all students across World'].astype('str')
        attempts = self.df.iloc[row]['First name\'s attempts (Attempts x 100 / Total Questions)'].astype('str')
        avg_attempts = self.df.iloc[row]['Average attempts of all students across the Worl'].astype('str')
        accuracy = self.df.iloc[row]['First name\'s Accuracy ( Corrects x 100 /Attempts )'].astype('str')
        avg_accuracy = self.df.iloc[row]['Average accuracy of all students across the World'].astype('str')
        #returning to main function
        return avg, median, mode, attempts, avg_attempts, accuracy, avg_accuracy

#function to show numerical value above the bars in the graph
def addlabels(x,y):
    for i in range(len(x)):
        plt.text(i, y[i], y[i], ha='center')

#function to show the header in first page
def head(pdf, f_name, l_name, reg_no, round):
    pdf.set_font('Arial', 'B', size=8)
    pdf.cell(190, 4, txt='Round ' + round +' - Enhanced Score Report: ' + f_name + " " + l_name, ln=1)
    pdf.cell(190, 4, txt='Reg Number: ' + reg_no, ln=1)

#defining the structure of the section 1 table (column headers)
def first_table_struture(pdf):
    pdf.cell(10, 5, ln=1)
    pdf.set_font('Arial', style='B', size=8)
    pdf.set_text_color(255, 255, 255)

    #row1
    pdf.cell(26, 4)
    pdf.cell(18, 4, txt="Question", border=1, align='C', fill=True)
    pdf.cell(20, 4, txt="Attempt", border=1, align='C', fill=True)
    pdf.cell(22, 4, txt=f_name + "\'s", border=1, align='C', fill=True)
    pdf.cell(18, 4, txt="Correct", border=1, align='C', fill=True)
    pdf.cell(20, 4, txt="", border=1, align='C', fill=True)
    pdf.cell(18, 4, txt="Score if", border=1, align='C', fill=True)
    pdf.cell(22, 4, txt=f_name + "\'s", border=1, align='C', fill=True, ln=1)

    #row2
    pdf.cell(26, 4)
    pdf.cell(18, 4, txt="No.", border=1, align='C', fill=True)
    pdf.cell(20, 4, txt="Status", border=1, align='C', fill=True)
    pdf.cell(22, 4, txt="Choice", border=1, align='C', fill=True)
    pdf.cell(18, 4, txt="Answer", border=1, align='C', fill=True)
    pdf.cell(20, 4, txt="Outcome", border=1, align='C', fill=True)
    pdf.cell(18, 4, txt="Correct", border=1, align='C', fill=True)
    pdf.cell(22, 4, txt="Score", border=1, align='C', fill=True, ln=1)

#defining the structure of the table for section 2 (column headers)
def second_table_structure(pdf):
    pdf.cell(10, 5, ln=1)
    pdf.set_font('Arial', style='B', size=8)
    pdf.set_text_color(255, 255, 255)

    #row1
    pdf.cell(15, 4, txt="", border=1, align='C', fill=True)
    pdf.cell(18, 4, txt="", border=1, align='C', fill=True)
    pdf.cell(18, 4, txt="", border=1, align='C', fill=True)
    pdf.cell(14, 4, txt="", border=1, align='C', fill=True)
    pdf.cell(18, 4, txt="", border=1, align='C', fill=True)
    pdf.cell(18, 4, txt="", border=1, align='C', fill=True)
    pdf.cell(25, 4, txt="", border=1, align='C', fill=True)
    pdf.cell(24, 4, txt="% of students", border=1, align='C', fill=True)
    pdf.cell(24, 4, txt="% of students", border=1, align='C', fill=True)
    pdf.cell(16, 4, txt="", border=1, align='C', fill=True, ln=1)

    #row2
    pdf.cell(15, 4, txt="", border=1, align='C', fill=True)
    pdf.cell(18, 4, txt="", border=1, align='C', fill=True)
    pdf.cell(18, 4, txt="", border=1, align='C', fill=True)
    pdf.cell(14, 4, txt="", border=1, align='C', fill=True)
    pdf.cell(18, 4, txt="", border=1, align='C', fill=True)
    pdf.cell(18, 4, txt="", border=1, align='C', fill=True)
    pdf.cell(25, 4, txt="% of students", border=1, align='C', fill=True)
    pdf.cell(24, 4, txt="(from those who", border=1, align='C', fill=True)
    pdf.cell(24, 4, txt="(from those who", border=1, align='C', fill=True)
    pdf.cell(16, 4, txt="World", border=1, align='C', fill=True, ln=1)

    #row3
    pdf.cell(15, 4, txt="", border=1, align='C', fill=True)
    pdf.cell(18, 4, txt="", border=1, align='C', fill=True)
    pdf.cell(18, 4, txt="", border=1, align='C', fill=True)
    pdf.cell(14, 4, txt="", border=1, align='C', fill=True)
    pdf.cell(18, 4, txt="", border=1, align='C', fill=True)
    pdf.cell(18, 4, txt="", border=1, align='C', fill=True)
    pdf.cell(25, 4, txt="across the world", border=1, align='C', fill=True)
    pdf.cell(24, 4, txt="attempted this)", border=1, align='C', fill=True)
    pdf.cell(24, 4, txt="attempted this)", border=1, align='C', fill=True)
    pdf.cell(16, 4, txt="Average", border=1, align='C', fill=True, ln=1)

    #row4
    pdf.cell(15, 4, txt="Question", border=1, align='C', fill=True)
    pdf.cell(18, 4, txt="Attempt", border=1, align='C', fill=True)
    pdf.cell(18, 4, txt=f_name + "\'s", border=1, align='C', fill=True)
    pdf.cell(14, 4, txt="Correct", border=1, align='C', fill=True)
    pdf.cell(18, 4, txt="", border=1, align='C', fill=True)
    pdf.cell(18, 4, txt=f_name + "\'s", border=1, align='C', fill=True)
    pdf.cell(25, 4, txt="who attempted", border=1, align='C', fill=True)
    pdf.cell(24, 4, txt="who got it", border=1, align='C', fill=True)
    pdf.cell(24, 4, txt="who got it", border=1, align='C', fill=True)
    pdf.cell(16, 4, txt="in this", border=1, align='C', fill=True, ln=1)

    #row5
    pdf.cell(15, 4, txt="No.", border=1, align='C', fill=True)
    pdf.cell(18, 4, txt="Status", border=1, align='C', fill=True)
    pdf.cell(18, 4, txt="Choice", border=1, align='C', fill=True)
    pdf.cell(14, 4, txt="Answer", border=1, align='C', fill=True)
    pdf.cell(18, 4, txt="Outcome", border=1, align='C', fill=True)
    pdf.cell(18, 4, txt="Score", border=1, align='C', fill=True)
    pdf.cell(25, 4, txt="this question", border=1, align='C', fill=True)
    pdf.cell(24, 4, txt="correct", border=1, align='C', fill=True)
    pdf.cell(24, 4, txt="incorrect", border=1, align='C', fill=True)
    pdf.cell(16, 4, txt="question", border=1, align='C', fill=True, ln=1)

#main function
if __name__ == "__main__":
    #name of the excel file to read data from
    s = student("Dummy Data for final assignment.xlsx")
    no_students, no_ques, photos_ls = s.preprocess()
    y_pos = 0
    #creating separate folder for storing report cards
    try:
       os.mkdir("Reports")
    except:
       pass

    #loop to create pdfs of results for different students
    for i in range(no_students):
        pdf = FPDF()
        pdf.add_page()
        #adding background
        pdf.image('back.jpg', x=0, y=0, w=210, h=300)
        photo, reg_no, f_name, l_name, dob, gender, grade, school, city, country, date, comment, round = s.student_details(i, photos_ls, no_ques)
        #adding header
        head(pdf, f_name, l_name, reg_no, round)
        pdf.set_font('Arial', style='B', size=12)
        pdf.cell(190, 7, txt="INTERNATIONAL MATHS OLYMPIAD CHALLENGE", ln=1, align='C')
        pdf.image('logo.jpg', x=75, y=25, w=60, h=30)

        pdf.cell(10, 30, ln=1)
        pdf.cell(190, 7, txt="Round " + round + " performance of " + f_name + " " + l_name, ln=1, align='C')
        pdf.image('./Student Photos/' + photo, x=165, y=23, w=35, h=35)
        pdf.cell(50, 5, txt="", ln=1)
        pdf.set_font('Arial', style='B', size=10)

        #firt 2 tables giving students details
        pdf.cell(12)
        pdf.cell(40, 6, txt="Grade ", border=1)
        pdf.set_font('Arial', style='', size=10)
        pdf.cell(40, 6, txt=" " + grade, border=1)
        pdf.cell(5)
        pdf.set_font('Arial', style='B', size=10)
        pdf.cell(40, 6, txt="Registration No ", border=1)
        pdf.set_font('Arial', style='', size=10)
        pdf.cell(40, 6, txt=" " + reg_no, ln=1, border=1)
        pdf.cell(12)
        pdf.set_font('Arial', style='B', size=10)
        pdf.cell(40, 6, txt="School ", border=1)
        pdf.set_font('Arial', style='', size=10)
        pdf.cell(40, 6, txt=" " + school, border=1)
        pdf.cell(5)
        pdf.set_font('Arial', style='B', size=10)
        pdf.cell(40, 6, txt="Gender ", border=1)
        pdf.set_font('Arial', style='', size=10)
        pdf.cell(40, 6, txt=" " + gender, ln=1, border=1)
        pdf.cell(12)
        pdf.set_font('Arial', style='B', size=10)
        pdf.cell(40, 6, txt="City of Residence ", border=1)
        pdf.set_font('Arial', style='', size=10)
        pdf.cell(40, 6, txt=" " + city, border=1)
        pdf.cell(5)
        pdf.set_font('Arial', style='B', size=10)
        pdf.cell(40, 6, txt="DOB ", border=1)
        pdf.set_font('Arial', style='', size=10)
        pdf.cell(40, 6, txt=" " + dob.strftime('%d/%m/%Y'), ln=1, border=1)
        pdf.cell(12)
        pdf.set_font('Arial', style='B', size=10)
        pdf.cell(40, 6, txt="Country of Residence ", border=1)
        pdf.set_font('Arial', style='', size=10)
        pdf.cell(40, 6, txt=" " + country, border=1)
        pdf.cell(5)
        pdf.set_font('Arial', style='B', size=10)
        pdf.cell(40, 6, txt="Date of Test ", border=1)
        pdf.set_font('Arial', style='', size=10)
        pdf.cell(40, 6, txt=" " + date, ln=1, border=1)

        pdf.set_font('Arial', style='BI', size=12)
        pdf.cell(10, 5, ln=1)
        pdf.cell(10, 5, ln=1)
        pdf.cell(190, 7, txt='Section 1', align='C', ln=1)
        pdf.set_font('Arial', style='', size=10)
        pdf.cell(190, 7, txt='This section describes ' + f_name + '\'s performance v/s the Test in Grade ' + grade, align='C', ln=1)

        #creating column heads of the table in section 1
        first_table_struture(pdf)

        #entering data into the table
        pdf.set_font('Arial', style='', size=8)
        pdf.set_text_color(0, 0, 0)
        start = i * no_ques
        end = start + no_ques
        total = 0
        for j in range(start, end):
          ques_no, student_answer, correct_answer, student_marks, max_marks, outcome = s.student_marks(j)
          pdf.cell(26, 5)
          pdf.cell(18, 5, txt=ques_no, border=1, align='C')
          if student_answer == 'Unattempted':
            attempt = 'Unattempted'
            student_answer = ''
          else:
            attempt = 'Attempted'
          pdf.cell(20, 5, txt=attempt, border=1, align='C')
          pdf.cell(22, 5, txt=student_answer, border=1, align='C')
          pdf.cell(18, 5, txt=correct_answer, border=1, align='C')
          pdf.cell(20, 5, txt=outcome, border=1, align='C')
          pdf.cell(18, 5, txt=max_marks, border=1, align='C')
          pdf.cell(22, 5, txt=student_marks, border=1, align='C', ln=1)
          total = total + student_marks.astype('int')

        pdf.set_font('Arial', style='BI', size=10)
        pdf.cell(124)
        pdf.cell(40, 8, txt="Total Score: " + total.astype('str'), align='C')

        pdf.add_page()
        pdf.image('back.jpg', x=0, y=0, w=210, h=300)
        pdf.set_font('Arial', style='BI', size=12)
        pdf.cell(10, 5, ln=1)
        pdf.cell(190, 7, txt='Section 2', align='C', ln=1)
        pdf.set_font('Arial', style='', size=10)
        pdf.cell(190, 7, txt='This section describes ' + f_name + '\'s performance v/s the Rest of the World in Grade ' + grade, align='C', ln=1)

        #creating column heads of the table in section 2
        second_table_structure(pdf)

        #entering data into the table
        pdf.set_font('Arial', style='', size=7)
        pdf.set_text_color(0, 0, 0)
        start = i * no_ques
        end = start + no_ques
        for j in range(start, end):
          ques_no, st_answer, correct_answer, st_marks, _, outcome = s.student_marks(j)
          d1, d2, d3, d4 = s.esr_report(j)
          pdf.cell(15, 4, txt=ques_no, border=1, align='C')
          if st_answer == 'Unattempted':
            attempt = 'Unattempted'
            st_answer = ''
          else:
            attempt = 'Attempted'
          pdf.cell(18, 4, txt=attempt, border=1, align='C')
          pdf.cell(18, 4, txt=st_answer, border=1, align='C')
          pdf.cell(14, 4, txt=correct_answer, border=1, align='C')
          pdf.cell(18, 4, txt=outcome, border=1, align='C')
          pdf.cell(18, 4, txt=st_marks, border=1, align='C')
          pdf.cell(25, 4, txt=d1, border=1, align='C')
          pdf.cell(24, 4, txt=d2, border=1, align='C')
          pdf.cell(24, 4, txt=d3, border=1, align='C')
          pdf.cell(16, 4, txt=d4, border=1, align='C', ln=1)

        pdf.set_font('Arial', style='', size=8)
        pdf.cell(2, 2, ln=1)
        pdf.multi_cell(190, 4, txt=comment)

        #creating the overview region
        pdf.set_font('Arial', style='B', size=8)
        pdf.cell(5, 5, ln=1)
        pdf.cell(20, 4, txt='Overview', ln=1)

        #creating the 3 tables
        avg, median, mode, attempts, avg_attempts, accuracy, avg_accuracy = s.overview(i, no_ques)
        pdf.set_font('Arial', style='', size=8)
        pdf.cell(42, 4, txt='Average score of all', border='LTR')
        pdf.cell(15, 4, txt='', border='LTR')
        pdf.cell(10, 4, txt='')
        pdf.cell(42, 4, txt=f_name + '\'s attempts', border='LTR')
        pdf.cell(15, 4, txt='', border='LTR')
        pdf.cell(10, 4, txt='')
        pdf.cell(42, 4, txt=f_name + '\'s Accuracy', border='LTR')
        pdf.cell(15, 4, txt='', border='LTR', ln=1)

        pdf.cell(42, 4, txt='students across the world', border='LR')
        pdf.cell(15, 4, txt=avg, border='LR')
        pdf.cell(10, 4, txt='')
        pdf.cell(42, 4, txt='(Attempts x 100/Total', border='LR')
        pdf.cell(15, 4, txt=attempts, border='LR')
        pdf.cell(10, 4, txt='')
        pdf.cell(42, 4, txt='(Corrects x 100/Attempts)', border='LR')
        pdf.cell(15, 4, txt=accuracy, border='LR', ln=1)

        pdf.cell(42, 4, txt='', border='LRB')
        pdf.cell(15, 4, txt='', border='LRB')
        pdf.cell(10, 4, txt='')
        pdf.cell(42, 4, txt='Questions)', border='LRB')
        pdf.cell(15, 4, txt='', border='LRB')
        pdf.cell(10, 4, txt='')
        pdf.cell(42, 4, txt='', border='LRB')
        pdf.cell(15, 4, txt='', border='LRB', ln=1)

        pdf.cell(42, 4, txt='Median score of all', border='LTR')
        pdf.cell(15, 4, txt='', border='LTR')
        pdf.cell(10, 4, txt='')
        pdf.cell(42, 4, txt='Average attempts of all', border='LTR')
        pdf.cell(15, 4, txt='', border='LTR')
        pdf.cell(10, 4, txt='')
        pdf.cell(42, 4, txt='Average accuracy of all', border='LTR')
        pdf.cell(15, 4, txt='', border='LTR', ln=1)

        pdf.cell(42, 4, txt='students across the world', border='LRB')
        pdf.cell(15, 4, txt=median, border='LRB')
        pdf.cell(10, 4, txt='')
        pdf.cell(42, 4, txt='students across the world', border='LRB')
        pdf.cell(15, 4, txt=avg_attempts, border='LRB')
        pdf.cell(10, 4, txt='')
        pdf.cell(42, 4, txt='students across the world', border='LRB')
        pdf.cell(15, 4, txt=avg_accuracy, border='LRB', ln=1)

        pdf.cell(42, 4, txt='Mode score of all', border='LTR')
        pdf.cell(15, 4, txt='', border='LTR', ln=1)

        pdf.cell(42, 4, txt='students across the World', border='LRB')
        pdf.cell(15, 4, txt=mode, border='LRB', ln=1)

        plt.style.use({'ytick.labelsize' : 12, 'xtick.labelsize' : 15, 'axes.titlesize' : 20, 'axes.labelsize' : 15})

        #creating first graph and saving as .jpg
        x=[f_name, 'Average', 'Median', 'Mode']
        y=[total, avg.astype('float'), median.astype('float'), mode.astype('float')]
        plt.figure(figsize=(5, 6))
        plt.bar(x, y, width = 0.7)
        addlabels(x, y)
        plt.ylabel("Score")
        plt.title("Comparision of Scores")
        plt.savefig('score.jpg')

        #creating second graph and saving as .jpg
        x=[f_name, 'World']
        y=[attempts.astype('float'), avg_attempts.astype('float')]
        plt.figure(figsize=(6, 6))
        plt.bar(x, y, width = 0.5)
        addlabels(x, y)
        plt.ylabel("Attempts(%)")
        plt.title("Comparision of Attempts(%)")
        plt.savefig('attempts.jpg')

        #creating third graph and saving as .jpg
        x=[f_name, 'World']
        y=[accuracy.astype('float'), avg_accuracy.astype('float')]
        plt.figure(figsize=(5, 6))
        plt.bar(x, y, width = 0.5)
        addlabels(x, y)
        plt.ylabel("Accuracy(%)")
        plt.title("Comparision of Accuracy(%)")
        plt.savefig('accuracy.jpg')

        #showing the graphs in the pdf
        y_pos = pdf.get_y()
        pdf.image('score.jpg', x=10, y=y_pos + 5, w=57, h=70)
        pdf.image('attempts.jpg', x=77, y=y_pos + 5, w=57, h=70)
        pdf.image('accuracy.jpg', x=144, y=y_pos + 5, w=57, h=70)

        #removing the graphs from the storage
        os.remove('score.jpg')
        os.remove('attempts.jpg')
        os.remove('accuracy.jpg')

        #saving the pdf
        pdf.output("./Reports/" + f_name + " " + l_name + " (" + reg_no + ").pdf")
