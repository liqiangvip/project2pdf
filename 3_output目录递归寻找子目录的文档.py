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

def word2PDF(wordFile, pdfFile):
    print(f'转换pdf: {wordFile}')
    w = gencache.EnsureDispatch('Word.Application')
    doc = w.Documents.Open(wordFile, ReadOnly=1)
    doc.ExportAsFixedFormat(pdfFile,
            constants.wdExportFormatPDF,
            Item=constants.wdExportDocumentWithMarkup,
            CreateBookmarks=constants.wdExportCreateHeadingBookmarks)
    w.Quit(constants.wdDoNotSaveChanges)

def findFilesInOutputDir(output_dir):
    for stu_dir in os.listdir(output_dir):
        if(stu_dir[0] == '.'):  # 排除隐藏文件
            continue
        for root, dirs, files in os.walk(os.path.join(output_dir, stu_dir)):  
            for filepath in files:
                if filepath.endswith(('.doc', '.docx')):
                    wordFile = os.path.join(root, filepath)
                    pdfFile = os.path.join(output_dir, stu_dir) + '.pdf'
                    word2PDF(wordFile, pdfFile)
                    time.sleep(0.2)
                elif filepath.endswith('.pdf'):
                    srcFileName = os.path.join(root, filepath)
                    dstfileName = os.path.join(output_dir, stu_dir) + '.pdf'
                    if os.path.exists(dstfileName):
                        dstfileName = os.path.join(output_dir, stu_dir) + str(randint(1,999)) + '.pdf'
                    shutil.copy(srcFileName, dstfileName)
        if os.path.isdir(os.path.join(output_dir, stu_dir)):
            print('删除 ...',os.path.join(output_dir, stu_dir))
            shutil.rmtree(os.path.join(output_dir, stu_dir))

def main():
    global extract_dir, total_count
    for f in os.listdir():
        if os.path.isdir(f):
            if(f[0] == '.'):  # 排除隐藏文件
                continue
            if f.endswith('_pdf'):
                findFilesInOutputDir(os.path.abspath(f))
                time.sleep(1)
    print('Stage2: DONE!')

main()