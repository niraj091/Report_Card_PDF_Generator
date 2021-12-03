import os
import jinja2
import random
import pdfkit
import pandas as pd
import matplotlib.pyplot as plt




path = os.getcwd()
path_wkhtmltopdf =os.path.join(path,'dependencies/wkhtmltopdf.exe')
config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)

class Student:
    def __init__(self,filename):
        self.filename = filename
        template_loader = jinja2.FileSystemLoader(searchpath="./")
        template_env = jinja2.Environment(loader=template_loader)
        self.template_file = "dependencies/report.html"
        
        self.template = template_env.get_template(self.template_file)

        if self.filename.endswith('.xlsx') or self.filename.endswith('.xls'):
            self.data = pd.read_excel(filename)
        elif self.filename.endswith('.csv'):
            self.data = pd.read_csv(filename)
        else:
            raise TypeError('Invalid File Extension')


        if self.data.columns.str.match('Unnamed').any():
            header_row = 0
            self.data.columns=self.data.iloc[header_row]
            self.data=self.data.drop(header_row)
            self.data=self.data.reset_index(drop=True)
        self.data=self.data.fillna('')
        self.registered=self.data.groupby(['Registration Number'],as_index=False).first()
        self.registered = self.registered['Registration Number'].to_list()
        self.group_data = self.data.groupby('Registration Number')

    def report(self,reg_no):
        self.section1=[]
        self.section2=[]
        self.param = []
        self.total_marks = 0
       
        
        if reg_no in self.registered:
            self.grouped_data=self.group_data.get_group(reg_no)

            for i in range(len(self.grouped_data)):
                self.data_dict=self.grouped_data.iloc[i].to_dict()
                self.report_data1 = self.grouped_data.iloc[i,[13,16,14,15,16,17,18]]
                self.section1.append(self.report_data1)
                self.report_data2 = self.grouped_data.iloc[i,[13,16,14,15,16,18,20,21,22,23]]
                self.section2.append(self.report_data2)
            
                self.total_marks = self.total_marks+int(self.section1[i][6])
            
            
           

            
            
            
            for j in range(len(self.section1)):
                if self.section1[j][1] !='Unattempted':
                    self.section1[j][1]='Attempted'
                if self.section2[j][1] !='Unattempted':
                    self.section2[j][1]='Attempted'
                if self.section1[j][2] == 'nan':
                    self.section1[j][2]= ''
                if self.section2[j][2] == 'nan':
                    self.section2[j][2]= ''
            for k in range(len(self.section2)):
                self.section2[k][6]=round(self.section2[k][6]*100,2)
                self.section2[k][7]=round(self.section2[k][7]*100,2)
                self.section2[k][8]=round(self.section2[k][8]*100,2)
            
            self.performance_plot()

            self.output_text = self.template.render(
            name=self.data_dict['Full Name '],
            reg_num=self.data_dict['Registration Number'],
            school=self.data_dict['Name of School '],
            date=self.data_dict['Date and time of test'],
            gender=self.data_dict['Gender'],
            grade=self.data_dict['Grade '],
            dob=self.data_dict['Date of Birth '],
            city=self.data_dict['City of Residence'],
            country=self.data_dict['Country of Residence'],
            round = self.data_dict['Round'],
            final_result=self.data_dict['Qualification'],
            data1 = self.section1,
            data2 = self.section2,
            total = self.total_marks,
            percentile=self.grouped_data.iloc[0][24],
            average_score = self.average_score,
            median = self.median,
            mode = self.mode,
            attempts = self.attempts,
            average_attempts = self.average_attempts,
            accuracy = self.accuracy,
            average_accuracy = self.average_accuracy,
            figure1=os.path.join(path,'dependencies/images/scores.png'),
            figure2=os.path.join(path,'dependencies/images/attempts.png'),
            figure3=os.path.join(path,'dependencies/images/accuracy.png'),
            path=path,
            )
        
            html_path = 'dependencies/'+self.data_dict['Full Name ']+'.html'
            html_file = open(html_path, 'w',encoding='utf-8')
            html_file.write(self.output_text)
            html_file.close()
            print("Generating PDF Report...")
            pdf_path = self.data_dict['Full Name ']+'('+str(reg_no)+')'+'.pdf' 
            self.html2pdf(html_path,pdf_path)
            os.remove(os.path.join(path,html_path))
            os.remove(os.path.join(path,'dependencies/images/scores.png'))
            os.remove(os.path.join(path,'dependencies/images/accuracy.png'))
            os.remove(os.path.join(path,'dependencies/images/attempts.png'))
            
            return 'PDF generated successfully.'
        return 'Student does not exist in the Database.'

    def performance_plot(self):
        self.name = self.data_dict['First Name ']
        self.average_score = self.grouped_data.iloc[0][24]
        self.median = self.grouped_data.iloc[0][25]
        self.mode = self.grouped_data.iloc[0][26]
        self.attempts = self.grouped_data.iloc[0][27]
        self.average_attempts = self.grouped_data.iloc[0][28]
        self.accuracy = round(self.grouped_data.iloc[0][29]*100,2)
        self.average_accuracy = round(self.grouped_data.iloc[0][30]*100,2)
        
   

        
        self.x1=[self.name,'Average','Median','Mode']
        self.y1=[self.total_marks,self.average_score,self.median,self.mode]
        for self.x1,self.y1 in zip(self.x1,self.y1):
            rgb=(random.random(),random.random(),random.random())
            plt.bar(self.x1,self.y1,color=[rgb],width=0.5)
        plt.ylim(0,100)
        plt.ylabel('SCORES',fontweight='bold')
        plt.title('COMPARISON OF SCORES',fontweight='bold')
        plt.savefig(os.path.join(path,'dependencies/images/scores.png'))
        plt.close()

        self.x2=[self.name,'World']
        self.y2=[self.attempts,self.average_attempts]
       
        for self.x2,self.y2 in zip(self.x2,self.y2):
            rgb=(random.random(),random.random(),random.random())
            plt.bar(self.x2,self.y2,color=[rgb],width=0.5)
        plt.ylim(0,100)
        plt.ylabel('ATTEMPTS(%)',fontweight='bold')
        plt.title('COMPARISON OF ATTEMPTS (%)',fontweight='bold')
        plt.savefig(os.path.join(path,'dependencies/images/attempts.png'))
        plt.close()

        self.x3=[self.name,'World']
        self.y3=[self.accuracy,self.average_accuracy]
        
        for self.x3,self.y3 in zip(self.x3,self.y3):
            rgb=(random.random(),random.random(),random.random())
            
            plt.bar(self.x3,self.y3,color=[rgb],width=0.5)
        plt.ylim(0,100)
        plt.ylabel('ACCURACY (%)',fontweight='bold')
        plt.title('COMPARISON OF ACCURACY (%)',fontweight='bold')
        
        plt.savefig(os.path.join(path,'dependencies/images/accuracy.png'))
        plt.close()


    def html2pdf(self,html_path, pdf_path):
        options = {
            'page-width':'240',
            'page-height':'290',
            'margin-top': '0in',
            'margin-right': '0in',
            'margin-bottom': '0in',
            'margin-left': '0in',
            'encoding': "UTF-8",
            'no-outline': None,
            'enable-local-file-access': None
        }
        with open(html_path,encoding='utf-8') as f:
            pdfkit.from_file(f, pdf_path, options=options,configuration=config)
    

if __name__=='__main__':
    while True:
        try:
            file_path=input('Enter file name or file path or "Q" to STOP:')
            if file_path == "Q":
                break
            regi_no = int(input('Enter Registration ID:'))
            s1 = Student(os.path.join(path,file_path))
            s1.report(regi_no)
        except Exception:
            print('Something went wrong.\nTry Again')
        
    




    




