# -*- coding:gb2312 -*-
import time
import logging
import shutil
import random
import os
import sys
import subprocess

import pythoncom
#get username
import getpass

#get pid, 2016/8/9 18:45:00
import psutil
import string

#get new testLog, 2016/8/10 14:51:29
import datetime

#打印异常，2016/10/18
import traceback

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#作用-读取文件到列表。使用方法:arrYear = ReadArrayFromFile("c:\years.txt")
#fileName-文件名
#返回数组列表
def ReadArrayFromFile(fileName):
  try:
    linelist=[]
    with open(fileName,'r') as f:
      for line in f.readlines():
          linestr = line.strip() #删除字符串开头、结尾空格，也可指定待删除字符
          linelist.append(linestr)

          #linestrlist = linestr.split("\n")
          #linelist = map(int,linestrlist)# 方法一
      f.close()
      return linelist  #返回列表
  except Exception, e:
    exstr = traceback.format_exc()
    print( "ReadArrayFromFile(), error!" + '\n' + exstr)
    pass  #空语句 do nothing
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#def main():
if __name__ == "__main__":
 try:
  #获取当前的路径
  currentPath = os.path.dirname(os.path.realpath(sys.argv[0] )) +'\\' #类似D:/jenkins/xx/
  if not os.path.exists(currentPath):
    print('程序退出，根路径不存在：' + currentPath)
    sys.exit(1)  #currentPath路径不存在：退出
  
  fileName=currentPath +"IpSeting.txt" #参数文件
  if not os.path.exists(fileName):
    print('IpSeting.txt参数文件不存在，程序退出！' + currentPath)
    sys.exit(1)  #currentPath路径不存在：退出
    
  with open(fileName,'r') as f: #获取secureCRT安装路径和待运行的软件名
      for line in f.readlines():
          linestr = line.rstrip() #删除字符串开头、结尾空格，也可指定待删除字符
          if "securecrtPath=" in linestr:   #1, secureCRT安装路径
             securecrtPath=linestr.split("=")[1]
             print("secureCRT安装路径--" +securecrtPath)
          if "softName=" in linestr:   #2, 设备SSH登录IP
             softName=currentPath +linestr.split("=")[1]
             print("待运行的软件--" +softName)
      f.close()
  
  if False == "SecureCRT.exe" in linestr:   #验证1, secureCRT设置
        print("在IpSeting.txt文件中CRT路径设置错误，程序退出！" +securecrtPath)
  
  if not os.path.exists(softName): #验证2, softName设置
    print("在IpSeting.txt文件设置的vbs路径不存在，程序退出！" +softName)
    sys.exit(1)  #softName路径不存在：退出
  
  softRun=securecrtPath +" /SCRIPT " + softName #防路径空格，程序位置与名称
  #os.popen(softRun).read() #阻塞执行exe程序，不能处理空格
  #retcode = subprocess.call(softRun) #阻塞执行exe程序
  subprocess.Popen(softRun) #非阻塞
  
  print("程序启动")
  
  #time.sleep(1)
  sys.exit(0)
 except Exception, e:
    exstr = traceback.format_exc()
    print( "startCheck.exe, error!" + '\n' + exstr)
    pass  #空语句 do nothing