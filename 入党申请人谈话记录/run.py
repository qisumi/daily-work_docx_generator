import pandas as pd
from docxtpl import DocxTemplate

source_data = pd.read_excel("source.xlsx")
template_doc = DocxTemplate("template.docx")
for index,person in source_data.iterrows():
    print("processing: ",index," ",person['姓名'])
    new_doc = template_doc
    data_docx = {
        'name':person['姓名'],
        'gender':person['性别'],
        'ethnic':person['民族'],
        'grade':int(person['一卡通号']/10000%100),
        'class_no':person['学号'][-3],
        'work':person['职务'],
        'birthday':"{}年{}月{}日".
            format(person["出生日期"].year,
                   person["出生日期"].month,
                   person["出生日期"].day),
        'apply_date':"{}年{}月{}日".
            format(person["申请入党时间"][0:4],
                   person["申请入党时间"][5:7],
                   person["申请入党时间"][8:10]),
        'addr':person["家庭住址"]
    }
    new_doc.render(data_docx)
    new_doc.save('output/'+person['姓名']+'.docx')