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

#��ӡ�쳣��2016/10/18
import traceback

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#����-��ȡ�ļ����б�ʹ�÷���:arrYear = ReadArrayFromFile("c:\years.txt")
#fileName-�ļ���
#���������б�
def ReadArrayFromFile(fileName):
  try:
    linelist=[]
    with open(fileName,'r') as f:
      for line in f.readlines():
          linestr = line.strip() #ɾ���ַ�����ͷ����β�ո�Ҳ��ָ����ɾ���ַ�
          linelist.append(linestr)

          #linestrlist = linestr.split("\n")
          #linelist = map(int,linestrlist)# ����һ
      f.close()
      return linelist  #�����б�
  except Exception, e:
    exstr = traceback.format_exc()
    print( "ReadArrayFromFile(), error!" + '\n' + exstr)
    pass  #����� do nothing
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#def main():
if __name__ == "__main__":
 try:
  #��ȡ��ǰ��·��
  currentPath = os.path.dirname(os.path.realpath(sys.argv[0] )) +'\\' #����D:/jenkins/xx/
  if not os.path.exists(currentPath):
    print('�����˳�����·�������ڣ�' + currentPath)
    sys.exit(1)  #currentPath·�������ڣ��˳�
  
  fileName=currentPath +"IpSeting.txt" #�����ļ�
  if not os.path.exists(fileName):
    print('IpSeting.txt�����ļ������ڣ������˳���' + currentPath)
    sys.exit(1)  #currentPath·�������ڣ��˳�
    
  with open(fileName,'r') as f: #��ȡsecureCRT��װ·���ʹ����е������
      for line in f.readlines():
          linestr = line.rstrip() #ɾ���ַ�����ͷ����β�ո�Ҳ��ָ����ɾ���ַ�
          if "securecrtPath=" in linestr:   #1, secureCRT��װ·��
             securecrtPath=linestr.split("=")[1]
             print("secureCRT��װ·��--" +securecrtPath)
          if "softName=" in linestr:   #2, �豸SSH��¼IP
             softName=currentPath +linestr.split("=")[1]
             print("�����е����--" +softName)
      f.close()
  
  if False == "SecureCRT.exe" in linestr:   #��֤1, secureCRT����
        print("��IpSeting.txt�ļ���CRT·�����ô��󣬳����˳���" +securecrtPath)
  
  if not os.path.exists(softName): #��֤2, softName����
    print("��IpSeting.txt�ļ����õ�vbs·�������ڣ������˳���" +softName)
    sys.exit(1)  #softName·�������ڣ��˳�
  
  softRun=securecrtPath +" /SCRIPT " + softName #��·���ո񣬳���λ��������
  #os.popen(softRun).read() #����ִ��exe���򣬲��ܴ���ո�
  #retcode = subprocess.call(softRun) #����ִ��exe����
  subprocess.Popen(softRun) #������
  
  print("��������")
  
  #time.sleep(1)
  sys.exit(0)
 except Exception, e:
    exstr = traceback.format_exc()
    print( "startCheck.exe, error!" + '\n' + exstr)
    pass  #����� do nothing