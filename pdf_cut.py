from pytesseract import * #pip install pytesseract
import os
import PyPDF2
from tkinter import *
import tkinter.filedialog
from pdf2image import convert_from_path
from PIL import Image


def setUp(month):
    global MONTH
    MONTH = month

#문자열 -> 텍스트파일 개별 저장
def strToTxt(txtName, outText):
    with open(txtName + '.txt', 'w', encoding='utf-8') as f:
        f.write(outText)

#영수증 첨부지인가?
def isReceipt(outText):
    if '비용코드' in outText or '영수증 첨부지' in outText or '지출비목' in outText or '사용내역' in outText:
        return True
    else:
        return False

#비목코드인가?
def get_cl(outText):
    outText=outText.replace(" ","")
    if '보수' in outText :
        return '보수'
    elif '일반수용비' in outText:
        return '일반수용비'
    elif '공공요금및제세' in outText:
        return '공공요금및제세'
    elif '특근매식비' in outText:
        return '특근매식비'
    elif '임차료' in outText:
        return '임차료'
    elif '재료비' in outText:
        return '재료비'
    elif '복리후생비' in outText:
        return '복리후생비'
    elif '국내여비' in outText:
        return '국내여비'
    elif '사업추진비' in outText:
        return '사업추진비'
    elif '월정직책금' in outText:
        return '월정직책금'
    elif '포상금' in outText:
        return '포상금'
    else:
        return ""


#텍스트로 되어 있는 페이지번호를 INT로 추출(998페이지까지 추출 가능)
def extract_page_num(fileName):
    if int(fileName[-7])>=1: ##100이상
        pageNum = (int(float(fileName[-7]))*100)+(int(float(fileName[-6])) * 10) + int(float(fileName[-5]))
    elif int(fileName[-6])>=1 : ##10이상
        pageNum = (int(float(fileName[-6])) * 10) + int(float(fileName[-5]))
    else:
        pageNum =int(float(fileName[-5]))
    pageNum = pageNum - 1

    return pageNum

#지출금액 추출
def extract_price(outText):
    price=""
    w = outText.find("액")
    w = w + 2
    string2 = outText[w:w + 60]
    print("스트링", string2)
    string2 = string2.replace("\n", "")
    string2 = string2.replace(" ", "")
    string2 = string2.replace(".",",")
    string2 = string2.replace("ㅣ", "")
    string2 = string2.replace("|", "")
    for i in range(0, len(string2)):
        if string2[i].isdigit() or string2[i] == ',' or string2[i]=='-':
            price = price + string2[i]
        else:
            break
    price=price[0:9]
    return price

#파일이름 만들기 ('월_순서_지출금액_비목')
def make_file_name(s_cl,s_fileNum,month,price,save_root):
    s_fileName =month+ "_" + str(s_fileNum) + "_" + str(price) +"_"+ str(s_cl)+".pdf"
    save_name = save_root+'/' + s_fileName
    return save_name

def extract_tree(in_file, out_file, start_num, last_num):
    with open(in_file, 'rb') as infp:
        # Read the document that contains the tree (in its first page)
        reader = PyPDF2.PdfFileReader(infp)
        writer = PyPDF2.PdfFileWriter()
        for i in range(start_num,last_num):
            writer.addPage(reader.getPage(i))
        with open(out_file, 'wb') as outfp:
            writer.write(outfp)

#이미지 크롭
def image_crop(save_root,image_dir) :
    try:
        directory = save_root + "/cropped_image"
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print('Error: Creating directory. ' + directory)
        return

    file_list = os.listdir(image_dir)
    for fname in file_list:
        print(fname)
        image1 = Image.open(image_dir+"/"+fname)
        croppedImage = image1.crop((50, 150, 1500, 470))
        croppedImage.save(directory +"/"+ fname)
    return directory

# 원본 pdf 파일 경로 묻기
def get_pdf_root():
    root = Tk().withdraw()
    pdf_root = tkinter.filedialog.askopenfilename(initialdir="/", title="PDF 파일 업로드", filetypes={("all files", "*.pdf")})
    print(pdf_root)
    return pdf_root

# 저장 폴더 경로 묻기
def get_save_root():
    root = Tk().withdraw()
    save_folder = tkinter.filedialog.askdirectory(title="저장 폴더 선택");
    print(save_folder)
    return save_folder

def pdf_to_img(pdf_root,save_root):
    try:
        directory = save_root + "/pdf2image"
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print('Error: Creating directory. ' + directory)

    pages = convert_from_path(pdf_root,fmt="jpeg")

    for i, page in enumerate(pages):
        pagenum = None
        if i < 9:
            pagenum = "00" + str(i + 1)
        elif i >= 9 and i < 99:
            pagenum = "0" + str(i + 1)
        elif i >= 99:
            pagenum = str(i + 1)
        page.save(
            directory+"/페이지" + pagenum + ".jpg",
            "JPEG")

    return directory

# def split_pdf(pdf_root,save_root):
def split_pdf(month,pdf_root,save_root):
    # pdf_root=get_pdf_root()
    # save_root=get_save_root()

    #전역변수 Month 설정
    setUp(month)
    global startPageNum, s_fileNum,s_fileName, s_fileNum, save,lastPageNum,price
    startPageNum = 10000
    s_fileNum=0
    s_fileName=MONTH
    price=""
    lastFileName=""

    # # 이미지 변환(전처리)
    ori_image_root=pdf_to_img(pdf_root,save_root)
    #crop_image_root=ori_image_root
    crop_image_root=image_crop(save_root,ori_image_root)

    # OCR 추출 작업 메인
    for root, dirs, files in os.walk(os.path.dirname(crop_image_root+"/")):
        for fileName in files:
            #OCR 추출 작업
            fullPath = os.path.join(root, fileName)
            img = Image.open(fullPath)
            outText = image_to_string(img, lang='kor', config='--psm 1 -c preserve_interword_spaces=1')
            # outText = image_to_string(img, lang='kor+eng', config='--psm 1 -c preserve_interword_spaces=1')
            print('+++ OCT Extract Result +++')
            print('Extract FileName ->>> : ', fileName, ' : <<<-\n\n',outText)
            if isReceipt(outText):
                # 시작페이지>=10000이면 처음 추출
                if startPageNum >= 10000:
                    startPageNum = extract_page_num(fileName)
                    s_cl = get_cl(outText)
                    if "액" in outText:
                        price = extract_price(outText)
                else:
                    lastPageNum = extract_page_num(fileName) ##다음 영수증첨부지가 있는 페이지
                    # 이미지파일은 1부터 시작하고, PDF 페이지는 0부터 시작
                    # extract_page_num은 PDF 인덱스 추출(-1 연산)
                    print ("파일 이름 ",fileName)
                    print(">>>>>>Start Page Num : ", startPageNum, ">>>>>>Last Page Num : ", lastPageNum)
                    s_fileNum = int(s_fileNum) + 1
                    save = make_file_name(s_cl,s_fileNum, MONTH, price,save_root)
                    extract_tree(pdf_root, save, startPageNum, lastPageNum)
                    startPageNum = lastPageNum

                    s_cl = get_cl(outText)
                    if "액" in outText:
                        price = extract_price(outText)
                    else:
                        price=""

            lastFileName=fileName

        lastPageNum = extract_page_num(lastFileName)
        s_fileNum = int(s_fileNum) + 1
        save = make_file_name(s_cl,s_fileNum, MONTH, price,save_root)
        extract_tree(pdf_root, save, startPageNum, lastPageNum+1)


#plit_pdf("4",get_pdf_root(),get_save_root())
#pdf_to_img(get_pdf_root(),get_save_root())