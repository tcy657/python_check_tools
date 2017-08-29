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
    set f = fso.CreateTextFile(filePath, True) '第二个参数表示同名文件存在时是否覆盖
    'f.Write("写入内容")
    'f.WriteLine("写入内容并换行")
    'f.WriteBlankLines(3) '写入三个空白行（相当于在文本编辑器中按三次回车）
    f.Write(log)
    f.Close()  'close是必要的,不要省
    set f = nothing
    set fso = Nothing
    writeFile = True
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function displayArray_L1(binArray)  '显示数组内容
    	 Dim str_1
    	 str_1=""
    	 For i3=0 To UBound(binArray,1)
        str_1=str_1 & binArray(i3) &Chr(13) '循环遍历数组，并输出数组值
       Next
    	 MsgBox "Array:"  & Chr(13) & str_1
    	 displayArray = True
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function displayArray_L2(binArray)  '显示数组内容,i行j列
    	 Dim str_1
    	 str_1="Array:"
    	 For i3=0 To UBound(binArray,1)'遍历行，第一维
    	   For j3=0 To UBound(binArray,2)'遍历列，第二维
          str_1=str_1 & "[" & i3 & ", " & j3 & "]- " & binArray(i3,j3) '循环遍历数组，并输出数组值
         Next
        str_1=str_1 & Chr(13) '一列处理完成，换行
       Next
    	 MsgBox str_1
    	 'writeFile checkResult, str_1
    	 displayArray_L2 = True
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function KillExcelProcess()
    on error resume Next
    CreateObject("WScript.Shell").Run "taskkill /f /im EXCEL.EXE ",0,True
    'kill所有的wps excel进程
    CreateObject("WScript.Shell").Run "taskkill /f /im et.exe ",0,True

    'CreateObject("WScript.Shell").Run "taskkill /f /im Wscript.exe"
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'打开FHAIP程序完整性检查记录
'fileName——带路径的文件名
'logStr——记录的内容
Function writeLogFile(fileNameStr, logStr)
    'KillExcelProcess
	Set binFileObject=CreateObject("Scripting.FileSystemObject")
    dim fLog
    If True = binFileObject.fileExists(fileNameStr) Then   '是否存在
      Set objFile = binFileObject.GetFile(fileNameStr)
       if objfile.Size >= 10000000 Then '文件大于10M？
         Set fLog = binFileObject.OpenTextFile(fileNameStr, 2, false) '第二个参数2表示重写
       else
	     Set fLog = binFileObject.OpenTextFile(fileNameStr, 8, false) '8 表示追加
	   end if
    Else
	  Set myfile = binFileObject.CreateTextFile(fileNameStr, true) '第二个参数表示目标文件存在时是否覆盖,不存在则创建
	  myfile.Close
	  Set fLog = binFileObject.OpenTextFile(fileNameStr, 8, false) '第二个参数8 表示追加
    End If

    fLog.WriteLine(Now & ": " & logStr)
	fLog.Close '退出文件

  Set binFileObject = Nothing   '释放文件操作对象
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'memory检查结果记录
'fileName——带路径的文件名
'logStr——记录的内容
Function writeResultFile(fileNameStr_1, logStr_1)
    'KillExcelProcess
	Set binFileObject_1=CreateObject("Scripting.FileSystemObject")
    dim fLog_1
    If True = binFileObject_1.fileExists(fileNameStr_1) Then   '是否存在
      Set fLog_1 = binFileObject_1.OpenTextFile(fileNameStr_1, 8, false) '第二个参数2表示重写，如果是 8 表示追加
    Else
	  Set myfile_1 = binFileObject_1.CreateTextFile(fileNameStr_1, true) '第二个参数表示目标文件存在时是否覆盖,不存在则创建
	  myfile_1.Close
	  Set fLog_1 = binFileObject_1.OpenTextFile(fileNameStr_1, 8, false) '第二个参数8 表示追加
    End If

    fLog_1.WriteLine(logStr_1) '不打印时间
	fLog_1.Close '退出文件

  Set binFileObject_1 = Nothing   '释放文件操作对象
End Function
''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'####################################################################  
'判断操作系统为中文或英文
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
'在脚本之间共享变量，以及相互调用函数。
'使用方法：Include  "libGlobal.vbs"
Sub Include(sInstFile) 
Dim oFSO, f, s 
Set oFSO = CreateObject("Scripting.FileSystemObject") 
Set f = oFSO.OpenTextFile(sInstFile) 
s = f.ReadAll 
f.Close 
ExecuteGlobal s 
End Sub 