# -*- coding: utf-8 -*-
"""
Created on Sat Jan 13 14:12:29 2018

@author: liqiangvip
"""
import os, os.path, sys, shutil, zipfile, time, chardet
if sys.platform == 'win32':
    from win32com.client import Dispatch, constants, gencache
from unrar import rarfile
from pathlib import Path
from random import randint

ZIP_FILENAME_UTF8_FLAG = 0x800

def decodeZipFileName(filename):
    '''
    对乱码的文件名进行解码还原
    '''
    try:
        #使用cp437对文件名进行解码还原
        filename = filename.encode('cp437')
        filename = filename.decode("gbk")
    except:
        #如果已被正确识别为utf8编码时则不需再编码
#        filename = filename.decode('utf-8')
        pass# 解压调用
    return filename

def unzip_file2(zfile_path, unzip_dir, encoding='gbk'):
    zf = zipfile.ZipFile(zfile_path, 'r')
    if not os.path.exists(unzip_dir):
        os.makedirs(unzip_dir)
    for file_info in zf.infolist():
        filename = file_info.filename
        if file_info.flag_bits & ZIP_FILENAME_UTF8_FLAG == 0:
            filename_bytes = filename.encode('437')
            guessed_encoding = chardet.detect(filename_bytes)['encoding'] or encoding
            filename = filename_bytes.decode(guessed_encoding, 'replace')
        if file_info.is_dir():
            os.mkdir(os.path.join(unzip_dir, filename))
            continue
        output_filename = os.path.join(unzip_dir, filename)
        output_file_dir = os.path.dirname(output_filename)
        if not os.path.exists(output_file_dir):
            os.makedirs(output_file_dir)
        with open(output_filename, 'wb') as output_file:
            shutil.copyfileobj(zf.open(file_info.filename), output_file)
    zf.close()

def unzip_file(zfile_path, unzip_dir):
    '''
    解压ZIP文件,基本能正确识别乱码文件名/目录名
    '''
    zf = zipfile.ZipFile(zfile_path, 'r')
    if not os.path.exists(unzip_dir):
        os.makedirs(unzip_dir)
    for file_info in zf.infolist():
        if file_info.is_dir():
            os.mkdir(os.path.join(unzip_dir, file_info.filename))
            continue
        filename = decodeZipFileName(file_info.filename)
        output_filename = os.path.join(unzip_dir, filename)
        output_file_dir = os.path.dirname(output_filename)
        if not os.path.exists(output_file_dir):
            os.makedirs(output_file_dir)
        with open(output_filename, 'wb') as output_file:
            shutil.copyfileobj(zf.open(file_info.filename), output_file)
    zf.close()
#    os.remove(zfile_path)
                                    
def unrar_file(rfile_path, unrar_dir):
    '''
    解压rar文件
    '''
    unrarfile = rarfile.RarFile(rfile_path)  #这里写入的是需要解压的文件，别忘了加路径
    unrarfile.extractall(path=unrar_dir)  #这里写入的是你想要解压到的文件夹

def word2PDF(wordFile, pdfFile):
    print(f'转换pdf: {wordFile}')
    w = gencache.EnsureDispatch('Word.Application')
    doc = w.Documents.Open(wordFile, ReadOnly=1)
    doc.ExportAsFixedFormat(pdfFile,
            constants.wdExportFormatPDF,
            Item=constants.wdExportDocumentWithMarkup,
            CreateBookmarks=constants.wdExportCreateHeadingBookmarks)
    w.Quit(constants.wdDoNotSaveChanges)

def processOutputDir(output_dir):
     # RAR文件解压没问题
    rarFiles = [fn for fn in os.listdir(output_dir) if fn.endswith(('.rar', '.RAR'))]
    print(rarFiles)
    for rarFile in rarFiles:
        rfile_path = os.path.join(output_dir ,os.path.basename(rarFile))
        unrar_dir = os.path.join(output_dir, rarFile.rsplit('.')[0])
        unrar_file(rfile_path, unrar_dir)
        time.sleep(0.2)

    # zip解压部分会有乱码问题
    zipFiles = [fn for fn in os.listdir(output_dir) if fn.endswith(('.zip', '.ZIP'))]
    print(zipFiles)
    for zipFile in zipFiles:
        zfile_path = os.path.join(output_dir ,os.path.basename(zipFile))
        unzip_dir = zfile_path.rsplit('.')[0]
        unzip_file(zfile_path, unzip_dir)
        time.sleep(0.2)

    wordFiles = [fn for fn in os.listdir(output_dir) if fn.endswith(('.doc','.docx'))]
    print(wordFiles)
    for wordFile in wordFiles:
        wordFile = os.path.join(output_dir, os.path.basename(wordFile))
        print(wordFile)
        index = wordFile.rfind('.')
        if index ==-1:
            continue
        pdfFile = wordFile[:index] + '.pdf'
        if os.path.exists(pdfFile):
            pdfFile = wordFile[:index] + str(randint(1,999)) + '.pdf'
        word2PDF(wordFile, pdfFile)
        time.sleep(0.2)

def clearOutputDir(output_dir):
    delFiles = [fn for fn in os.listdir(output_dir) if fn.endswith(('.doc','.docx','rar', 'zip'))]
    for delFile in delFiles:
        delFile = os.path.join(output_dir ,os.path.basename(delFile))
        os.remove(delFile)
        time.sleep(0.1)

def main():
    global extract_dir, total_count
    for f in os.listdir():
        if os.path.isdir(f):
            if(f[0] == '.'):  # 排除隐藏文件
                continue
            if f.endswith('_pdf'):
                processOutputDir(os.path.abspath(f))
                time.sleep(1)
                clearOutputDir(os.path.abspath(f))
    print('Stage2: DONE!')

main()