#$language = "VBScript"
#$interface = "1.0"

'created by cytao@fiberhome.com, 2016/5/17 14:19:30
'on error resume Next

crt.Screen.Synchronous = True

Dim g_objTab,g_szSkippedTabs,cmdUsername,cmdPassword
Set g_objTab = crt.GetScriptTab
g_szSkippedTabs = ""
g_objTab.Screen.Synchronous = True

'dbgVar=True
dbgVar=False
cmdUsername="none"  'telnet RCU 的用户名
cmdPassword="fiberhome" 'telnet RCU的密码

Dim arrDeviceIP(), arrCMD(), ipNumber
countOk=0 '正常站点数
countBad=0 '异常站点数
ipNumber = 0 '记录IP站点的个数

Dim Board() '首先定义一个一维动态数组
'槽位号+盘名+R1X+CRC统计
ReDim Board(50,3) '重新定义为二维数组,51行4列。

Sub Main
    On Error Resume Next
	     If crt.Dialog.MessageBox(_
        "check NP board after 3s?" & vbcrlf & vbcrlf , _
        "check NP board", _
        vbyesno) <> vbyes then exit Sub

    '第一步，获取pingSetings文件参数
    Dim currentPath, objFSO, MyArray, binArray, R86xYesNo
       Set objFSO = CreateObject("Scripting.FileSystemObject")
      currentPath = createobject("Scripting.FileSystemObject").GetFolder(".").Path + "\"      '获取当前的路径
    '--------------------------------------------------------------------
      szData =  now '记录本次操作时间
      szData = Replace(szData, "/", "-")  '中划线
      szData = Replace(szData, " ", "_")  '下划线
      szData = Replace(szData, ":", "-")  '中划线

      fileDate= currentPath&"NPversion_checkLog" & ".log"
      writeLogFile fileDate, "check starts, NPversion_checkLog-" & fileDate
      writeLogFile fileDate, "Option Time:" & szData

	Set objFile = objFSO.OpenTextFile(currentPath&"IpSeting.txt", 1)
    d = 0
    resultPath="none"  '存放路径
    Do Until objFile.AtEndOfStream
    	line=objFile.ReadLine
       if instr(line,"ip=") then  '1, 设备SSH登录IP
          MyArray = Split(line, "=", -1, 1)
          line = MyArray(1)
          Redim Preserve arrDeviceIP(ipNumber)  '添加IP
          arrDeviceIP(ipNumber) = line
          ipNumber = ipNumber + 1
       end if
       if instr(line,"username=") then  '2, 登录RCU的用户名
          MyArray = Split(line, "=", -1, 1)
          line = MyArray(1)
          cmdUsername=line
          username_flag= 1	'标志置1，正常
       end if
       if instr(line,"userpwd=") then  '3, 登录RCU的密码
          MyArray = Split(line, "=", -1, 1)
          line = MyArray(1)
          cmdPassword=line
        userpwd_flag= 1  '标志置1，正常
       end if
       if instr(line,"resultPath=") then  '4, 获取结果保存路径
         MyArray = Split(line, "=", -1, 1)
         line = MyArray(1)
	      if right(line, 1) ="\" then '获最后一个字符，判断是否为”\“
	        resultPath= line
	      else '加个符号，防止使用人员忘记加了
	        resultPath= line	 + "\"
         end if
        end if

  Loop

  if UCase(resultPath) <> "NONE" and objFSO.FolderExists(resultPath) then
      filePath=resultPath&"NPversion_checkResult" & szData & ".csv"
      writeLogFile fileDate, "resultPath exists! --"& resultPath
  else
       filePath=currentPath&"NPversion_checkResult" & szData & ".csv"
       writeLogFile fileDate, "resultPath is none or resultPath donot exist, value is--"& resultPath
  end if

  if (objFSO.fileexists(filePath)) then   '判断result.csv文件是否存在
  	  objFSO.deletefile(filePath)         '删除result.csv文件
  End if
  writeLogFile fileDate, "NPversion_check result --" & filePath
  str_1="NeIP, boardName, slot, NpVersion, checkTime--" & now '写入报头
  writeResultFile filePath, str_1 '写入报头

  writeLogFile fileDate, "kill excel and csv program for reading csv file"
  KillExcelProcess
  crt.Sleep 200

    objFile.Close
    Set objFSO =Nothing
    For ii_1=0 To UBound(arrDeviceIP)-LBound(arrDeviceIP) '1th#,
        writeLogFile fileDate, "NPversion_check NE IP-" & arrDeviceIP(ii_1)
	   If crt.Session.Connected Then crt.Session.Disconnect        ' #如果有已建立的连接则断开连接。
       cmd = "/ssh2 /L root" & " /PASSWORD root" & " /C 3DES " & arrDeviceIP(ii_1)
       numVar=0
       do while numVar < 3  '3次重试
          err.clear

    	   if numVar=0 then
    	     crt.Session.ConnectInTab cmd   '第一次在tab中连接
    	   else
    	     crt.Session.Connect cmd		  '第2,3次在本窗口中连接
    	   end if
    	   'crt.sleep 3000
         If Err.Number <> 0 Then  '登录不正常
             numVar=numVar+1
    	     if numVar > 3 or numVar =3 then
    	       writeLogFile fileDate, "Exit for 3 times fail!-" & arrDeviceIP(ii_1)
               if g_szSkippedTabs = "" then
                   g_szSkippedTabs = crt.Window.Caption  & vbcrlf '索引号为nIndex
               else
                   g_szSkippedTabs = g_szSkippedTabs & "," & crt.Window.Caption & vbcrlf
               end if
			      str_1=crt.Window.Caption & ", /,  because of not connected， NE cannot be checked"
           	      writeResultFile filePath, str_1 '保存到文件
    	       exit do '结束本站巡检
    	     End if
    	     'crt.sleep 3000
          Else   '登录成功
    	   numVar =3

           Dim nIndex
           nIndex = 1 '2th#,
               Set objCurrentTab = crt.GetTab(nIndex)
               objCurrentTab.Activate
               ' Skip tabs that aren't connected
               if objCurrentTab.Session.Connected = True then

       	         'do sth, end
				 R86xYesNo="none" 'yes-R86x, no-R845
       	          if ( "NONE" = UCase(cmdUsername) ) then '无用户名
                     SendExpect nIndex,"telnet 127.1 2650", "Password:"
	                 SendExpect nIndex, cmdPassword, ">"
	                 SendExpect nIndex,"en", "#"
                  Else  '有用户名
                     SendExpect nIndex,"telnet 127.1 2650", "Username:"
	                 SendExpect nIndex, cmdUsername, "Password:"
	                 SendExpect nIndex, cmdPassword, ">"
	                 SendExpect nIndex,"en", "#"
                  End if
       	          szData = CaptureOutputOfCommand(nIndex, "show tne board", "#")
				  writeLogFile fileDate, "show tne board --" & szData
				  'get R86x or R845?
				  scur1=0
				  scuo1=0
				  scup1=0
				  scuq1=0
				  R86xYesNo="none"
				  '86x设备巡检不统计SCU盘，凡带SCU字样的单盘都过滤。
				  scur1=instr(1,szData,"SCUR1")
				  scuo1=instr(1,szData,"SCUO")
				  '845设备巡检只统计SCU盘。
       	          scup1=instr(1,szData,"SCUP1")
				  scuq1=instr(1,szData,"SCUQ1")
				  if ( 0 = scup1 and 0 = scuq1 ) and (0 < scur1 or 0 < scuo1) then 'device is R86x
				     R86xYesNo="yes"
				  	 writeLogFile fileDate, "device is R86x"
				  elseif ( 0 < scup1 or 0 < scuq1 ) and (0 = scur1 and 0 = scuo1) then 'device is R845
 				     R86xYesNo="no"
					 writeLogFile fileDate,  "device is R845"
				  else
				     writeLogFile fileDate, "device type is not sure, default for R86x device "
					 R86xYesNo="yes"
					 'exit do '结束本站巡检
				  end if
       	          SendExpect nIndex,"exit", "root"

       	          'Step2: get board list
                   MyArray = Split(szData, vbcrlf, -1, 1)
                   j=0 '索引，盘名和槽位
				   writeLogFile fileDate, "read show-tne-board to get slot and boardName"
                   For ii_3=0 To UBound(MyArray,1)   '3th#,
       	                s1=instr(1,MyArray(ii_3),"0x380")
						'86x设备巡检不统计SCU盘，凡带SCU字样的单盘都过滤。
						scur1=instr(1,MyArray(ii_3),"SCUR1")
						scuo1=instr(1,MyArray(ii_3),"SCUO") 'SCUO1 or SCUO2
						'845设备巡检只统计SCU盘。
       	                scup1=instr(1,MyArray(ii_3),"SCUP1")
						scuq1=instr(1,MyArray(ii_3),"SCUQ1")
                         
						 scuBoardYesNo=False 'case1: SCUX board
						 scuBoardYesNo=0 < scur1 or 0 < scuo1 or 0 < scup1 or 0 < scuq1
						 
						 If s1 > 0 And "none" <> R86xYesNo and True =scuBoardYesNo Then 'case1: SCUX board
						     writeLogFile fileDate,  "board is R845/R86x SCUSX"
                         	 tmp=""  '替换连续空格
                         	 MyArray(ii_3)=Trim(MyArray(ii_3)) '去掉开始和尾部空格
                            MyArray(ii_3)=replace(MyArray(ii_3)," ",",")
                            for ii_4 = 1 to len(MyArray(ii_3))  '4th#,
                             i3=ii_4+1
                             c1=mid(MyArray(ii_3),ii_4,1)  '取当前字符
                             If i3< len(MyArray(ii_3)) Then
                               c2=mid(MyArray(ii_3),i3,1)
                               If c1<>"," Or c2<>"," Then
                                 tmp =tmp & c1
                               End If
                             Else
                             	tmp =tmp & c1
                             End If
                            Next  '4th#_Next,

                         	   MyArray(ii_3)=tmp

                         	 binArray=Split(MyArray(ii_3), ",", -1, 1)
                         	   Board(j,0)=binArray(0)  'slot
                         	   Board(j,1)=binArray(2)  'board name
                         	   j=j+1 'next
                         End If
						 
						 If s1 > 0 And "yes" = R86xYesNo and False =scuBoardYesNo Then 'case2: R86x Np board
                         	 writeLogFile fileDate,  "board is R86x Np"
							 tmp=""  '替换连续空格
                         	 MyArray(ii_3)=Trim(MyArray(ii_3)) '去掉开始和尾部空格
                            MyArray(ii_3)=replace(MyArray(ii_3)," ",",")
                            for ii_4 = 1 to len(MyArray(ii_3))  '4.1th#,
                             i3=ii_4+1
                             c1=mid(MyArray(ii_3),ii_4,1)  '取当前字符
                             If i3< len(MyArray(ii_3)) Then
                               c2=mid(MyArray(ii_3),i3,1)
                               If c1<>"," Or c2<>"," Then
                                 tmp =tmp & c1
                               End If
                             Else
                             	tmp =tmp & c1
                             End If
                            Next  '4.1th#_Next,

                         	   MyArray(ii_3)=tmp

                         	 binArray=Split(MyArray(ii_3), ",", -1, 1)
                         	   Board(j,0)=binArray(0)  'slot
                         	   Board(j,1)=binArray(2)  'board name
                         	   j=j+1 'next
                         End If
						 
                   Next   '3th#_Next,

                   writeLogFile fileDate, "telnet 10.26.0.x"
                   j=j-1 '多加了1
       	          For ii_5=0 To j   '5th#, 遍历行，获取槽位信息
       	            writeLogFile fileDate, "telnet 10.26.0." &  Board(ii_5,0)
					SendExpect nIndex,"telnet 10.26.0." & Board(ii_5,0), "VxWorks login: "
       	            SendExpect nIndex,"bmu852", "Password: "
       	            SendExpect nIndex,"aaaabbbb", "->"

       	            szData = CaptureOutputOfCommand(nIndex, "DbgGetDriverConfig", "->")  'case 1: 0xB1 or 0xB2
       	            writeLogFile fileDate, "case 1: 0xB1 or 0xB2 --" &szData
					If dbgVar Then MsgBox "readmii2(1,2)-" & szData End If
       	            s1=instr(1,szData,"npu_chip_ver              = 0xB1")  '0xB1
       	            s2=instr(1,szData,"npu_chip_ver              = 0xB2")   '0xB2
       	            If  s1>0 and 0=s2 Then '0xB1
       	            	Board(ii_5,2)="0xB1"
       	            elseIf  0=s1 and s2>0 Then '0xB2
       	            	Board(ii_5,2)="0xB2"
       	            else '防错,输出原始值
					    Board(ii_5,2)="error（type is not sure）"
					End If

					SendExpect nIndex,"logout", "root"   '退出单盘 

       	          Next  '5th#_Next,

       	          writeLogFile fileDate, "save check result: NeIP	boardName slot	NpVersion"
       	          str_1=""
           	      For ii_6=0 To j  '6th#,
                    str_1=crt.Window.Caption & "," &  Board(ii_6,1) & "," &  Board(ii_6,0) &  _
                          "," &  Board(ii_6,2)  'Chr(13)
                    writeResultFile filePath, str_1 '保存到文件
				    writeLogFile fileDate, str_1 '保存到日志
				  Next   '6th#_Next,
       	        'do sth, end
              End if
           '2th#_Next,
		 end if
		Loop
    Next    '1th#_Next,
    g_objTab.Activate

    if g_szSkippedTabs <> "" Then   '巡检结束
        g_szSkippedTabs = vbcrlf & vbcrlf & _
            "because of not connected, NE cannot be checked：" & _
            vbcrlf & vbtab & g_szSkippedTabs
    end If

    crt.Dialog.MessageBox _
        "check end!" & vbcrlf & _
        vbtab & g_szSkippedTabs,"check end!", BUTTON_OK
    'close secureCRT.exe
	CreateObject("WScript.Shell").Run "taskkill /f /im SecureCRT.exe",0,True
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function SendExpect(nIndex,szSend, szExpect)
    SendExpect=False
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
    	 'writeFile filePath, str_1
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
'CRC检查结果记录
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
