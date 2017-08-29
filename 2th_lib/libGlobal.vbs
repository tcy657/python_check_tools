'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function SendExpect(nIndex,szSend, szExpect)
    Set objCurrentTab = crt.GetTab(nIndex)
    timeVar=0
    if objCurrentTab.Session.Connected <> True then exit function
        
    Do Until timeVar > 2
      objCurrentTab.Screen.Send szSend & vbcr
      If objCurrentTab.Screen.WaitForString(szExpect, 10) <> True Then
        timeVar=timeVar+1
      else
        timeVar=3
      End If
    Loop

    SendExpect = True
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function CaptureOutputOfCommand(nIndex, szCommand, szPrompt)
   Set objCurrentTab = crt.GetTab(nIndex)
   if objCurrentTab.Session.Connected <> True then
       CaptureOutputOfCommand = "[ERROR: Not Connected.]"
       exit function
   end if
    
    timeVar=0
    Do Until timeVar > 2
     objCurrentTab.Screen.Send szCommand & vbcr
     objCurrentTab.Screen.WaitForString vbcr
     CaptureOutputOfCommand = objCurrentTab.Screen.ReadString(szPrompt,3)
     If CaptureOutputOfCommand &" " =" " Then
     	  timeVar=timeVar+1
     Else
         timeVar=3
     End If
    Loop
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function writeFile(filePath,  log)
    dim fso, f
    set fso =CreateObject("Scripting.FileSystemObject")
    set f = fso.CreateTextFile(filePath, True) '�ڶ���������ʾͬ���ļ�����ʱ�Ƿ񸲸�
    'f.Write("д������")
    'f.WriteLine("д�����ݲ�����")
    'f.WriteBlankLines(3) 'д�������հ��У��൱�����ı��༭���а����λس���
    f.Write(log)
    f.Close()  'close�Ǳ�Ҫ��,��Ҫʡ
    set f = nothing
    set fso = Nothing
    writeFile = True
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function displayArray_L1(binArray)  '��ʾ��������
    	 Dim str_1
    	 str_1=""
    	 For i3=0 To UBound(binArray,1)
        str_1=str_1 & binArray(i3) &Chr(13) 'ѭ���������飬���������ֵ
       Next
    	 MsgBox "Array:"  & Chr(13) & str_1
    	 displayArray = True
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function displayArray_L2(binArray)  '��ʾ��������,i��j��
    	 Dim str_1
    	 str_1="Array:"
    	 For i3=0 To UBound(binArray,1)'�����У���һά
    	   For j3=0 To UBound(binArray,2)'�����У��ڶ�ά
          str_1=str_1 & "[" & i3 & ", " & j3 & "]- " & binArray(i3,j3) 'ѭ���������飬���������ֵ
         Next
        str_1=str_1 & Chr(13) 'һ�д�����ɣ�����
       Next
    	 MsgBox str_1
    	 'writeFile checkResult, str_1
    	 displayArray_L2 = True
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function KillExcelProcess()
    on error resume Next
    CreateObject("WScript.Shell").Run "taskkill /f /im EXCEL.EXE ",0,True
    'kill���е�wps excel����
    CreateObject("WScript.Shell").Run "taskkill /f /im et.exe ",0,True

    'CreateObject("WScript.Shell").Run "taskkill /f /im Wscript.exe"
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'��FHAIP���������Լ���¼
'fileName������·�����ļ���
'logStr������¼������
Function writeLogFile(fileNameStr, logStr)
    'KillExcelProcess
	Set binFileObject=CreateObject("Scripting.FileSystemObject")
    dim fLog
    If True = binFileObject.fileExists(fileNameStr) Then   '�Ƿ����
      Set objFile = binFileObject.GetFile(fileNameStr)
       if objfile.Size >= 10000000 Then '�ļ�����10M��
         Set fLog = binFileObject.OpenTextFile(fileNameStr, 2, false) '�ڶ�������2��ʾ��д
       else
	     Set fLog = binFileObject.OpenTextFile(fileNameStr, 8, false) '8 ��ʾ׷��
	   end if
    Else
	  Set myfile = binFileObject.CreateTextFile(fileNameStr, true) '�ڶ���������ʾĿ���ļ�����ʱ�Ƿ񸲸�,�������򴴽�
	  myfile.Close
	  Set fLog = binFileObject.OpenTextFile(fileNameStr, 8, false) '�ڶ�������8 ��ʾ׷��
    End If

    fLog.WriteLine(Now & ": " & logStr)
	fLog.Close '�˳��ļ�

  Set binFileObject = Nothing   '�ͷ��ļ���������
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'memory�������¼
'fileName������·�����ļ���
'logStr������¼������
Function writeResultFile(fileNameStr_1, logStr_1)
    'KillExcelProcess
	Set binFileObject_1=CreateObject("Scripting.FileSystemObject")
    dim fLog_1
    If True = binFileObject_1.fileExists(fileNameStr_1) Then   '�Ƿ����
      Set fLog_1 = binFileObject_1.OpenTextFile(fileNameStr_1, 8, false) '�ڶ�������2��ʾ��д������� 8 ��ʾ׷��
    Else
	  Set myfile_1 = binFileObject_1.CreateTextFile(fileNameStr_1, true) '�ڶ���������ʾĿ���ļ�����ʱ�Ƿ񸲸�,�������򴴽�
	  myfile_1.Close
	  Set fLog_1 = binFileObject_1.OpenTextFile(fileNameStr_1, 8, false) '�ڶ�������8 ��ʾ׷��
    End If

    fLog_1.WriteLine(logStr_1) '����ӡʱ��
	fLog_1.Close '�˳��ļ�

  Set binFileObject_1 = Nothing   '�ͷ��ļ���������
End Function
''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'####################################################################  
'�жϲ���ϵͳΪ���Ļ�Ӣ��
Function language() 
 language="None"
 strComputer = "." 
 Set objWMIService = GetObject("winmgmts://" &strComputer &"/root/CIMV2") 
 Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem") 
 For Each objItem In colItems   
   language = objItem.OSLanguage 
   If language = "1033" Then 
     'Language = "EN" 
     language = "English" 
   elseif language = "2052" then 
     'Language = "CN" 
     language = "Chinese" 
   End If     
 Next 
End Function            
''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'�ڽű�֮�乲��������Լ��໥���ú�����
'ʹ�÷�����Include  "libGlobal.vbs"
Sub Include(sInstFile) 
Dim oFSO, f, s 
Set oFSO = CreateObject("Scripting.FileSystemObject") 
Set f = oFSO.OpenTextFile(sInstFile) 
s = f.ReadAll 
f.Close 
ExecuteGlobal s 
End Sub 