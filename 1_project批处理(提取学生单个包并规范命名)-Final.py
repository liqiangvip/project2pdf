# 批量转PDF.py
# Created by liqiang on 10/01/2018.
# coding=utf-8

import os, os.path, shutil, zipfile, time, sys
if sys.platform == 'win32':
    from win32com.client import Dispatch, constants, gencache


id2name = {}
name2id = {}
pdf_count = 0
word_count = 0
total_count = 0

def outputFile(filename, output_dir):
    global total_count
    print(f'输出文件:{filename}-->{output_dir}')
    shutil.copy(filename, output_dir+'/'+os.path.basename(filename))
    total_count += 1

def word2pdf(wordFileName):
    #print('转换word文件:', wordfilename)
    f_name, extend_name = os.path.splitext(wordFileName)
    if extend_name.lower() in ('.docx', '.doc'):
        index = wordFileName.rindex('.')
        pdfFileName = wordFileName[:index] + '.pdf'
        print(f'转换PDF: -->{pdfFileName}')
        pass

def word2PDF(wordFile, pdfFile):
    print(f'转换PDF: -->{pdfFile}')
    w = gencache.EnsureDispatch('Word.Application')
    doc = w.Documents.Open(wordFile, ReadOnly=1)
    doc.ExportAsFixedFormat(pdfFile,
            constants.wdExportFormatPDF,
            Item=constants.wdExportDocumentWithMarkup,
            CreateBookmarks=constants.wdExportCreateHeadingBookmarks)
    w.Quit(constants.wdDoNotSaveChanges)
    
def idStuNameAsFileName(stu_dir, oldfilename):
    stuName = os.path.basename(stu_dir)
    file_name, extend_name = oldfilename.rsplit('.', 1)
    newfileName = os.path.join(stu_dir, name2id[stuName]+stuName+'.'+extend_name.lower())
    print(f"文件改名: -->{newfileName}")
    if os.path.exists(newfileName):
        pass
    else:
        os.rename(os.path.join(stu_dir, oldfilename), newfileName)
    return newfileName
    
def stuNameAsFileName(stu_dir, oldfilename):
    stuName = os.path.basename(stu_dir)
    file_name, extend_name = oldfilename.rsplit('.', 1)
    newfileName = os.path.join(stu_dir, stuName+'.'+extend_name.lower())
    print(f"文件改名: -->{newfileName}")
    os.rename(os.path.join(stu_dir, oldfilename), newfileName)
    return newfileName    

def extract_all(zip_filename, extract_dir, filename_encoding='GBK'):
    zf = zipfile.ZipFile(zip_filename, 'r')
    for file_info in zf.infolist():
        filename = file_info.filename
        try:
            #使用cp437对文件名进行解码还原
            filename = filename.encode('cp437')
            filename = filename.decode("gbk")
        except:
            #如果已被正确识别为utf8编码时则不需再编码
            filename = filename.decode('utf-8')
            pass# 解压调用
        print('解压... 获得...', filename)
        output_filename = os.path.join(extract_dir, filename)
        output_file_dir = os.path.dirname(output_filename)
        if not os.path.exists(output_file_dir):
            os.makedirs(output_file_dir)
        with open(output_filename, 'wb') as output_file:
            shutil.copyfileobj(zf.open(file_info.filename), output_file)
    zf.close()
    #print(f'删除zip文件... ', zip_filename)
    #os.remove(zip_filename)

def loadStuInfo():
    global id2name, name2id
    with open('stuinfo.csv', encoding='utf-8') as fp:
        lines = [line.strip().split(',') for line in fp.readlines()]
        id2name = {k.strip():v.strip() for k,v in lines}
        name2id = {v.strip():k.strip() for k,v in lines}
    return 

def processStuHW(stu_dir):
    global output_path
    print(f'处理学生目录:{stu_dir}')
    for f in os.listdir(stu_dir):
        if os.path.isfile(os.path.join(stu_dir, f)):
            index = f.rfind('.')
            if index == -1:
                continue
            extend_name = f[index:]
            if(f[0] == '.'):  # 排除隐藏文件
                continue
            elif extend_name.lower() in ('.doc', '.docx'):
                if f.startswith('计算机科学与技术学院'):
                    continue
                else:
                    wordFileName = idStuNameAsFileName(stu_dir, f)
                    if sys.platform == 'win32':
#                        index = wordFileName.rindex('.')
#                        pdfFileName = wordFileName[:index]+ '.pdf'
#                        print(wordFileName, pdfFileName)
#                        word2PDF(wordFileName, pdfFileName)
#                        time.sleep(0.1)
#                        outputFile(pdfFileName, output_path)
                        word2pdf(wordFileName)
                        outputFile(wordFileName, output_path)
                    else:
                        word2pdf(wordFileName)
                        outputFile(wordFileName, output_path)
            elif extend_name.lower() in ('.pdf', '.zip', '.rar'):
                newfileName = idStuNameAsFileName(stu_dir, f) 
                outputFile(newfileName, output_path)
    return

def processSingleStuDir(path):
    '''
    对每个文件夹单独处理
    '''
    global name2id
    for f in os.listdir(path):
        if os.path.isdir(os.path.join(path, f)):
            if(f[0] == '.'):
                continue
            elif f in name2id:
                processStuHW(os.path.join(path, f))
            else:
                continue

def exactSingleStuZipFile(path):
    '''
    处理一次作业或者学习报告的文件夹,里面含有很多个学生的打包作业
    '''
    global output_path
    # 先解压所有子目录中的压缩文件
    for f in os.listdir(path):
        if os.path.isfile(os.path.join(path, f)):
            index = f.rfind('.')
            if index == -1:
                continue
            extend_name = f[index:]
            if(f[0] == '.'):  # 排除隐藏文件
                pass
            elif extend_name.lower() == '.zip':
                extract_all(os.path.join(path, f), os.path.join(path, f[:index]))
                time.sleep(0.1)

def processSingleProject(projectName):
    global output_path, total_count
    projectPath = os.path.join(os.getcwd(), projectName)
    output_path = projectPath+'_pdf'
    if not os.path.exists(output_path):
        os.mkdir(output_path)
    exactSingleStuZipFile(projectPath)
    time.sleep(5)
    processSingleStuDir(projectPath)
    print(f'一共输出文件: {total_count}')

def processOutputDir(wordfile_dir):
    wordFiles = [fn for fn in os.listdir(wordfile_dir) if fn.endswith(('.doc','.docx'))]
    print(wordFiles)
    for wordFile in wordFiles:
        wordFile = os.path.abspath(wordFile)
        index = wordFile.rindex('.')
        pdfFile = wordFile[:index] + '.pdf'
        word2PDF(wordFile, pdfFile)
        time.sleep(0.1)

def main():
    global extract_dir, total_count
    loadStuInfo()
    projectNames = ['第1次学习报告(以太网)','第2次学习报告(二层交换)', '第3次学习报告(IS-IS)',
                    'LAB-RIP', 'LAB-VLAN', 'LAB-STP', 'LAB-OSPF']
    for proj in projectNames:
        processSingleProject(proj)
    print('Stage1: DONE!')

main()
