from python_docx_replace import *
from docxtpl import *
from docx2pdf import convert
import os
import shutil
from PyPDF2 import PdfMerger,PdfReader
import glob
#context = {'ddd':'ssss'}
#doc.render(context)
#doc.save('new.docx')


#lang_ver = [pl:{name:'pl',template_path:'./template_pl',version:'polska'}]
en_version = {
'doc_types':[
{
'name':'1.Syllabus with Curricullum',
'doc_title':'Part 1/5 - Syllabus with curriculum',
'doc_title2':'Syllabus with curriculum'
},{
'name':'2.Trainers Manual',
'doc_title':'Part 2/5 - Materials for trainers',
'doc_title2':'Materials for trainers'
},{
'name':'3.Materials for Participants',
'doc_title':'Part 3/5 - Training materials for participants',
'doc_title2':'Training materials for participants'
},{
'name':'4.Training Evaluation Questionnaire',
'doc_title':'Part 4/5 - Training course evaluation questionnaire',
'doc_title2':'Training course evaluation questionnaire'
},{
'name':'5.Validation Tools',
'doc_title':'Part 5/5 - Tools for validation of learning outcomes',
'doc_title2':'Tools for validation of learning outcomes'
}],
'courses':[
{'en_title':'Course 1_Literacy','title':'Literacy','author':'MiA'},
{'en_title':'Course 4_Entrepreneurship','title':'Entrepreneurship','author':'Deinde'},
{'en_title':'Course 2_Digital Competences','title':"Digital Competences", 'author':'Inercia'},
{'en_title':'Course 3_Personal, social and learning to learn','title':'Personal, social and learning to learn','author':'ISC'}
],
'template_loc':'./template_en.docx', 'version':'English','sv':'EN'}
es_version = {
'doc_types':[
{
'name':'1.Syllabus with Curricullum',
'doc_title':'Parte 1/5 - Plan de estudios',
'doc_title2':'Plan de estudios'
},{
'name':'2.Trainers Manual',
'doc_title':'Parte 2/5 - Material para formadores (escenario de formación/manual del formador)',
'doc_title2':'Material para formadores (escenario de formación/manual del formador)'
},{
'name':'3.Materials for Participants',
'doc_title':'Parte 3/5 - Materiales de formación para los participantes',
'doc_title2':'Materiales de formación para los participantes'
},{
'name':'4.Training Evaluation Questionnaire',
'doc_title':'Parte 4/5 - Cuestionario de evaluación del curso de formación',
'doc_title2':'Cuestionario de evaluación del curso de formación'
},{
'name':'5.Validation Tools',
'doc_title':'Parte 5/5 - Herramientas para evaluar los resultados de la formación',
'doc_title2':'Herramientas para evaluar los resultados de la formación'
}],
'courses':[
{'en_title':'Course 2_Digital Competences','title':"Competencias digitales", 'author':'Inercia'},
{'en_title':'Course 3_Personal, social and learning to learn','title':'Competencias personales, sociales y de aprendizaje','author':'ISC'},
{'en_title':'Course 1_Literacy','title':'Comprensión y creación de información','author':'MiA'},
{'en_title':'Course 4_Entrepreneurship','title':'espíritu emprendedor','author':'Deinde'}
],
'template_loc':'./template_es.docx', 'version':'Español','sv':"ES"}
pl_version = {
'doc_types':[
{
'name':'1.Syllabus z programem nauczania',
'doc_title':'Część 1/5 – Syllabus z programem nauczania',
'doc_title2':'Syllabus z programem nauczania'
},{
'name':'2.Podręcznik trenera',
'doc_title':'Część 2/5 – Podręcznik trenera',
'doc_title2':'Podręcznik trenera'
},{
'name':'3.Materiały dla uczestników',
'doc_title':'Część 1/5 – Syllabus z programem nauczania',
'doc_title2':'Syllabus z programem nauczania'
},{
'name':'4.Kwestionariusz oceny kursu szkoleniowego',
'doc_title':'Część 4/5 – Kwestionariusz oceny kursu szkoleniowego',
'doc_title2':'Kwestionariusz oceny kursu szkoleniowego'
},{
'name':'5.Narzędzia walidacji efektów kształcenia',
'doc_title':'Część 5/5 – Narzędzia walidacji efektów kształcenia',
'doc_title2':'Narzędzia walidacji efektów kształcenia'
}],
'courses':[
{'title':"w zakresie kompetencji cyfrowych", 'author':'Inercia'},
{'title':'Kompetencje osobiste, społeczne i w zakresie umiejętności uczenia się','author':'ISC'},
{'title':'w zakresie rozumienia i tworzenia informacji','author':'MiA'},
{'title':'w zakresie przedsiębiorczości','author':'Deinde'}
],
'version':'polska','template_loc':'./template_pl.docx','sv':'PL'}

lang_versions=[en_version]




def gen_Pages():
    for lang_version in lang_versions:
        doc_loc = lang_version.get('template_loc')
        doc = DocxTemplate(doc_loc)
        version=lang_version.get('version')
        courses = lang_version.get('courses')
        sv = lang_version.get('sv')
        doc_types = lang_version.get('doc_types')
        
        for course in courses:
            en_title = course.get('en_title')
            title = course.get('title')
            author = course.get('author')
            isExist = os.path.exists("./"+en_title)
            if isExist==False:
                os.mkdir("./"+en_title)
            isExist2 = os.path.exists("./"+en_title+'/'+sv+'/templates')
            if isExist2==False:
                os.makedirs("./"+en_title+'/'+sv+'/templates')
            print('path is')
                
            for doc_type in doc_types:
                doc_name = doc_type.get('name')
                doc_title = doc_type.get('doc_title')
                doc_title2 = doc_type.get('doc_title2')
                doc_type = doc_type.get('type')
                
                context = {'title':title,'doc_title2':doc_title2,'doc_title':doc_title,'type':doc_type,'author':author,'version':version}
                doc.render(context)
                doc.save('./'+en_title+'/'+sv+'/templates/'+doc_name+'.docx')
                print(title,author,version)

#gen_Pages()
def move(src,dest):
    print(src,dest)
    allfiles = os.listdir(src)
    for file in allfiles:
        print(file)
        isExist = os.path.exists(dest)
        if isExist==False:
            os.makedirs(dest)
        if file.endswith('.pdf'):
            shutil.move(src+'/'+file, dest+'/'+file)
            
def convert2nd():       
    for lang_version in lang_versions:
        sv=lang_version.get('sv')
        courses = lang_version.get('courses')
    for course in courses:
        en_title = course.get('en_title')
        print(en_title)
        convert('./'+en_title+'/'+sv+'/templates/')
        src='./'+en_title+'/'+sv+'/templates/'
        dest='./'+en_title+'/'+sv+'/pdf2/'
        move(src,dest)
  

def convertWords():
    for lang_version in lang_versions:
        sv=lang_version.get('sv')
        courses = lang_version.get('courses')
    for course in courses:
        en_title = course.get('en_title')
        print(en_title)
        convert('./'+en_title+'/'+sv+'/word/')
        move('./'+en_title+'/'+sv+'/word','./'+en_title+'/'+sv+'/pdf3')


    
def moveWords():
    for lang_version in lang_versions:
        version=lang_version.get('version')
        courses = lang_version.get('courses')
    for course in courses:
        en_title=course.get('en_title')
        author = course.get('author')
        allfiles = os.listdir('./'+en_title+'/'+version+'/word')
        for file in allfiles:
            isExist = os.path.exists('./'+en_title+'/'+version+'/pdf3')
            if isExist==False:
                os.mkdir('./'+en_title+'/'+version+'/pdf3')
            if file.endswith('.pdf'):
                shutil.move('./'+en_title+'/'+version+'/word/'+file, './'+en_title+'/'+version+'/pdf3/'+file)    


def filtr(ex,seq):
   print(ex,seq)
def mergePdf():
    courses=en_version.get('courses')
    for course in courses:
        sv=en_version.get('sv')
        en_title=course.get('en_title')
        author = course.get('author')
        path = './'+en_title+'/'+sv+'/'
        covers_path = path+'covers/'
        second_path = path+'pdf2/'
        word_path = path+'pdf3/'
        back_path = path+'covers/back.pdf'
        isExist = os.path.exists(path+'/final')
        if isExist==False:
            os.mkdir(path+'/final')    
        words = os.listdir(word_path)
        second_pages=os.listdir(second_path)
        wordPdfs=[]
        for word in words:
            
            cover=covers_path+word[0]+'.pdf'
            back = covers_path+'/back.pdf'
            
            second_page=''
            for page in second_pages:
                if page.startswith(word[0]):
                    second_page=page
            wordPdfs.append([cover,second_path+second_page,word_path+word,back])
        #print(wordPdfs)
           
        for pdf in wordPdfs:
            merger = PdfMerger(strict=False)
            print(pdf)
            for doc in pdf:
                merger.append(fileobj=open('./'+doc, 'rb'))
            name=pdf[2].replace(word_path,'')
            merger.write(fileobj=open(path+'/final/'+name, 'wb'))
            merger.close()

        #merger.write(fileobj=open('./fina.pdf', 'wb'))
        #merger.close()







def mergePDF():
    for lang_version in lang_versions:
        version=lang_version.get('version')
        courses = lang_version.get('courses')
    for course in courses:
        en_title=course.get('en_title')
        author = course.get('author')
        covers_path = './'+en_title+'/'+version+'/covers/'
        second_path = en_title+'/'+version+'/pdf2/'
        word_path = en_title+'/'+version+'/pdf3/'
        back_path = en_title+'/'+version+'/covers/back.pdf'
        isExist = os.path.exists('./'+en_title+'/'+version+'/final')
        if isExist==False:
            os.mkdir('./'+en_title+'/'+version+'/final')    
        words = os.listdir(word_path)
        wordPdfs=[]
        for word in words:
            #wordPdfs.concat(word)
            cover=covers_path+word[0]+'.pdf'
            second_pages=os.listdir(second_path)
            second_page=''
            for page in second_pages:
                if page.startswith(word[0]):
                    second_page=(page)
            print(cover)
            back=back_path
            merger= PdfMerger
            merger.append(PdfReader(open(cover)))
            merger.append(second_page)
            merger.append(word)
            merger.append(back)
            merger.write('./'+en_title+'/'+version+'/end/'+pdfname)
            merger.close
#mergePDF()


#convert2nd() 
#convertWords()    
mergePdf()
