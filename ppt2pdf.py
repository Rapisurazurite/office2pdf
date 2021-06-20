# -*- coding: utf-8 -*-
import shutil

import comtypes.client
import os

from tqdm import tqdm


def init_powerpoint():
    powerPoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerPoint.Visible = 1
    return powerPoint


def ppt_to_pdf(powerPoint, inputFileName, outputFileName: str, formatType=32):
    outputFileName = os.path.splitext(outputFileName)[0]
    if outputFileName[-3:] != 'pdf':
        outputFileName += ".pdf"
    try:
        deck = powerPoint.Presentations.Open(inputFileName)
        deck.SaveAs(outputFileName, formatType)  # formatType = 32 for ppt to pdf
        deck.Close()
    except Exception as e:
        print("Input File Name" + inputFileName)
        print(e)


def convert_files_in_folder(powerPoint, folder):
    files = os.listdir(folder)  # 回指定文件夹包含的文件或文件夹名字的列表
    pdf_folder = os.path.join(folder, "PDF")
    pptFiles = [f for f in files if f.endswith((".ppt", ".pptx"))]  # 使用循环批量转换
    for pptFile in tqdm(pptFiles):
        pdfName = os.path.splitext(pptFile)[0] + ".pdf"
        if os.path.isfile(os.path.join(folder, pdfName)) or os.path.isfile(os.path.join(pdf_folder, pdfName)):
            continue
        fullPath = os.path.join(cwd, pptFile)  # 将多个路径组合后返回
        tqdm.write("Convert {}...".format(fullPath))
        ppt_to_pdf(powerPoint, fullPath, fullPath)


def moveFolder(folder: str):
    pdf_folder = os.path.join(folder, "PDF")
    if not os.path.exists(pdf_folder):
        os.mkdir(pdf_folder)
    files = os.listdir(folder)
    pdfFiles = [f for f in files if f.endswith(".pdf")]
    for pdfFile in pdfFiles:
        shutil.move(os.path.join(folder, pdfFile), os.path.join(pdf_folder, pdfFile))


if __name__ == "__main__":
    powerpoint = init_powerpoint()
    cwd = os.getcwd()  # 返回当前进程的目录
    convert_files_in_folder(powerpoint, cwd)
    powerpoint.Quit()
    moveFolder(cwd)
