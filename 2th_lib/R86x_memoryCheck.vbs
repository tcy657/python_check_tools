#$language = "VBScript"
#$interface = "1.0"

'created by cytao@fiberhome.com, 9:31 2017-8-25
'on error resume Next

crt.Screen.Synchronous = True

Dim g_objTab,g_szSkippedTabs,cmdUsername,cmdPassword
Set g_objTab = crt.GetScriptTab
g_szSkippedTabs = ""
g_objTab.Screen.Synchronous = True

dbgVar=True
'dbgVar=False
cmdUsername="none"  'telnet RCU 的用户名
cmdPassword="fiberhome" 'telnet RCU的密码

Dim arrDeviceIP(), arrCMD(), ipNumber
countOk=0 '正常站点数
countBad=0 '异常站点数
ipNumber = 0 '记录IP站点的个数


Dim Board() '首先定义一个一维动态数组
'槽位号+盘名+maxBlock+memory统计
ReDim Board(50,3) '重新定义为二维数组,51行4列。

dim MyArray  '临时存放

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Sub Main
    On Error Resume Next
      If crt.Dialog.MessageBox(_
        "check R86x NP/SCUxx memory after 3 seconds?" & vbcrlf & vbcrlf , _
        "check-Confirm", _
        vbyesno) <> vbyes then exit Sub
     
    '第一步，获取pingSetings文件参数
    Dim currentPath, objFSO
       Set objFSO = CreateObject("Scripting.FileSystemObject")     
       currentPath = createobject("Scripting.FileSystemObject").GetFolder(".").Path + "\"      '获取当前的路径
       Include currentPath& "libGlobal.vbs"
	   '创建log文件夹
	   if objFSO.FolderExists(currentPath &"log")<>True then
           objFSO.CreateFolder(currentPath &"log") '创建文件夹,目标文件夹的父文件夹必须存在
       end if
	   checkLog= currentPath &"log\86x_memCheck.log"

    '--------------------------------------------------------------------
      szData =  now '记录本次操作时间
      szData = Replace(szData, "/", "-")  '中划线
      szData = Replace(szData, " ", "_")  '下划线
      szData = Replace(szData, ":", "-")  '中划线
      languageResut=language '获取语言类型 
      
      writeLogFile checkLog, "86x_memCheckLog-" & checkLog
      writeLogFile checkLog, "optionTime:" & szData

	Set objFile = objFSO.OpenTextFile(currentPath&"\IpSeting.txt", 1)                                                                                                 
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
      checkResult=resultPath&"86x_memCheckResult" & szData & ".csv"
      writeLogFile checkLog, "resultPath exists--"& resultPath
  else
       checkResult=currentPath&"86x_memCheckResult" & szData & ".csv"
       writeLogFile checkLog, "resultPath is none or doesnot exist, value is--"& resultPath
  end if

  if (objFSO.fileexists(checkResult)) then   '判断result.csv文件是否存在
  	  objFSO.deletefile(checkResult)         '删除result.csv文件
  End if
  writeLogFile checkLog, "86x_mem check result-" & checkResult
  'str_1="站点名,单盘名称,槽位,maxBlock,memory统计结果, 本次巡检时间--" & now '写入报头
  str_1="NeName,BoardName,slot,maxBlock,freeMemory(byte), checkTime--" & now '写入报头
  writeResultFile checkResult, str_1 '写入报头
  	
  writeLogFile checkLog, "kill excel and wps that open csv file"
  KillExcelProcess  
  crt.Sleep 200
                                                                                                                
    objFile.Close    
    Set objFSO =Nothing     	 
    For ii_1=0 To UBound(arrDeviceIP)-LBound(arrDeviceIP) '1th#, 
        writeLogFile checkLog, "now ,we check 86x_mem IP-" & arrDeviceIP(ii_1)
	   If crt.Session.Connected Then crt.Session.Disconnect        ' #如果有已建立的连接则断开连接。
       cmd = "/ssh2 /ACCEPTHOSTKEYS /L root" & " /PASSWORD root" & " /C 3DES " & arrDeviceIP(ii_1)
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
    	       writeLogFile checkLog, "Exit for 3 times fail!-" & arrDeviceIP(ii_1)
               if g_szSkippedTabs = "" then
                   g_szSkippedTabs = crt.Window.Caption  & vbcrlf '索引号为nIndex
               else
                   g_szSkippedTabs = g_szSkippedTabs & "," & crt.Window.Caption & vbcrlf
               end if		
			      str_1=crt.Window.Caption & ", /,  ssh failed" 
           	      writeResultFile checkResult, str_1 '保存到文件
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
       	          SendExpect nIndex,"exit", "root" 
       	          
       	          'Step2: get board list
                   MyArray = Split(szData, vbcrlf, -1, 1)
                   j=0 '索引，盘名和槽位
				   writeLogFile checkLog, "read all lines, get NP and scu board slot number--"
                   For ii_3=0 To UBound(MyArray,1)   '3th#, 
       	                s1=instr(1,MyArray(ii_3),"0x380")  'memory巡检统计SCU/NP盘
                         If s1 > 0 Then
                         	 dim binArray
                         	 
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
                   Next   '3th#_Next, 
                   
                   writeLogFile checkLog, "telnet every NP/SCU board"
                   j=j-1 '多加了1
       	          For ii_5=0 To j   '5th#, 遍历行，获取槽位信息
       	            writeLogFile checkLog, "read all lines, telnet all NP/SCU--" &  Board(ii_5,0)
					SendExpect nIndex,"telnet 10.26.0." & Board(ii_5,0), "VxWorks login: " 
       	            SendExpect nIndex,"bmu852", "Password: " 
       	            SendExpect nIndex,"aaaabbbb", "->" 
       
       	            szData = CaptureOutputOfCommand(nIndex, "memShow", "->")  '判断1：memShow
       	            writeLogFile checkLog, "check: memShow - " &szData
       	            
                    MyArray = Split(szData, vbcrlf, -1, 1)
                    For ii_3=0 To UBound(MyArray,1)   '3th#, get "free" line
                       s1=instr(1,MyArray(ii_3),"free")
                       If s1 > 0 Then
                          freeStr=MyArray(ii_3)
                          exit for
                       end if
                    Next    
                               
                   writeLogFile checkLog, "free: memShow - " & freeStr  
                   tmp=""  '4th#, 替换连续空格
                   binString=Trim(freeStr) '去掉开始和尾部空格
                   binString=replace(binString," ",",")
                   for ii_4 = 1 to len(binString) 
                    i3=ii_4+1
                    c1=mid((binString),ii_4,1)  '取当前字符
                    If i3< len(binString) Then 
                      c2=mid(binString,i3,1) 
                      If c1<>"," Or c2<>"," Then
                        tmp =tmp & c1
                      End If  
                    Else 
                    	tmp =tmp & c1
                    End If       
                   Next  '4th#_Next, 
                   
                     freeStr=tmp
                      writeLogFile checkLog, "free(, replace ): memShow - " &freeStr
                   MyArray = Split(freeStr, ",", -1, 1)
                   MaxBlockNumber = CLng(MyArray(4)) 'MaxBlockNumber
                   freeBytesNumber = CLng(MyArray(1)) 'freeBytesNumber
                    writeLogFile checkLog, "2 numbers: memShow - " &MaxBlockNumber & "-" & freeBytesNumber
                   if MaxBlockNumber<200000 then 'normal, >=200K; error <200K
                       MaxBlockResult="error: value=" &MaxBlockNumber
                   else
                       MaxBlockResult="ok: value>200K"
                   end if 
                   
                   writeLogFile checkLog, "Result: memShow - " &MaxBlockResult & "-" & freeBytesNumber 
                    
       	            '槽位号+盘名+maxBlock+memory统计
       	            Board(ii_5,2)=MaxBlockResult  'maxBlock
                    Board(ii_5,3)=freeBytesNumber 'freeBytes
       	            
       	            SendExpect nIndex,"logout", "root"   '退出单盘

       	           Next  '5th#_Next, 
				   objCurrentTab.Screen.Send "exit" & vbcr    '退出会话
       
       	          writeLogFile checkLog, "save check result,NeName BoardName slot maxBlock freeMemory"
       	          str_1=""
           	      For ii_6=0 To j  '6th#, 
                    str_1=crt.Window.Caption & "," &  Board(ii_6,1) & "," &  Board(ii_6,0) &  _
                          "," &  Board(ii_6,2) & "," &  Board(ii_6,3) 'Chr(13)
                    writeResultFile checkResult, str_1 '保存到文件
				    writeLogFile checkLog, str_1 '保存到日志
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
            "because of failed connect, these checked NE without result: " & _
            vbcrlf & vbtab & g_szSkippedTabs
    end If
             
    crt.Dialog.MessageBox _
        "memory check finish!!" & _
        vbtab & g_szSkippedTabs & vbcrlf & _ 
		 "check finish, result file is at- " &checkResult,"check finish", BUTTON_OK

End Sub    
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