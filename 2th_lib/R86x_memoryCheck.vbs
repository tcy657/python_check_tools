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
cmdUsername="none"  'telnet RCU ���û���
cmdPassword="fiberhome" 'telnet RCU������

Dim arrDeviceIP(), arrCMD(), ipNumber
countOk=0 '����վ����
countBad=0 '�쳣վ����
ipNumber = 0 '��¼IPվ��ĸ���


Dim Board() '���ȶ���һ��һά��̬����
'��λ��+����+maxBlock+memoryͳ��
ReDim Board(50,3) '���¶���Ϊ��ά����,51��4�С�

dim MyArray  '��ʱ���

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Sub Main
    On Error Resume Next
      If crt.Dialog.MessageBox(_
        "check R86x NP/SCUxx memory after 3 seconds?" & vbcrlf & vbcrlf , _
        "check-Confirm", _
        vbyesno) <> vbyes then exit Sub
     
    '��һ������ȡpingSetings�ļ�����
    Dim currentPath, objFSO
       Set objFSO = CreateObject("Scripting.FileSystemObject")     
       currentPath = createobject("Scripting.FileSystemObject").GetFolder(".").Path + "\"      '��ȡ��ǰ��·��
       Include currentPath& "libGlobal.vbs"
	   '����log�ļ���
	   if objFSO.FolderExists(currentPath &"log")<>True then
           objFSO.CreateFolder(currentPath &"log") '�����ļ���,Ŀ���ļ��еĸ��ļ��б������
       end if
	   checkLog= currentPath &"log\86x_memCheck.log"

    '--------------------------------------------------------------------
      szData =  now '��¼���β���ʱ��
      szData = Replace(szData, "/", "-")  '�л���
      szData = Replace(szData, " ", "_")  '�»���
      szData = Replace(szData, ":", "-")  '�л���
      languageResut=language '��ȡ�������� 
      
      writeLogFile checkLog, "86x_memCheckLog-" & checkLog
      writeLogFile checkLog, "optionTime:" & szData

	Set objFile = objFSO.OpenTextFile(currentPath&"\IpSeting.txt", 1)                                                                                                 
    d = 0 
    resultPath="none"  '���·��
    Do Until objFile.AtEndOfStream 
    	line=objFile.ReadLine                       
       if instr(line,"ip=") then  '1, �豸SSH��¼IP
          MyArray = Split(line, "=", -1, 1) 
          line = MyArray(1)                                                                            
          Redim Preserve arrDeviceIP(ipNumber)  '���IP                                                                              
          arrDeviceIP(ipNumber) = line
          ipNumber = ipNumber + 1          
       end if                                                                                                          
       if instr(line,"username=") then  '2, ��¼RCU���û���
          MyArray = Split(line, "=", -1, 1)
          line = MyArray(1)
          cmdUsername=line
          username_flag= 1	'��־��1������
       end if
       if instr(line,"userpwd=") then  '3, ��¼RCU������
          MyArray = Split(line, "=", -1, 1)
          line = MyArray(1)
          cmdPassword=line
        userpwd_flag= 1  '��־��1������
       end if
       if instr(line,"resultPath=") then  '4, ��ȡ�������·��
         MyArray = Split(line, "=", -1, 1)
         line = MyArray(1)
	      if right(line, 1) ="\" then '�����һ���ַ����ж��Ƿ�Ϊ��\��
	        resultPath= line
	      else '�Ӹ����ţ���ֹʹ����Ա���Ǽ���
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

  if (objFSO.fileexists(checkResult)) then   '�ж�result.csv�ļ��Ƿ����
  	  objFSO.deletefile(checkResult)         'ɾ��result.csv�ļ�
  End if
  writeLogFile checkLog, "86x_mem check result-" & checkResult
  'str_1="վ����,��������,��λ,maxBlock,memoryͳ�ƽ��, ����Ѳ��ʱ��--" & now 'д�뱨ͷ
  str_1="NeName,BoardName,slot,maxBlock,freeMemory(byte), checkTime--" & now 'д�뱨ͷ
  writeResultFile checkResult, str_1 'д�뱨ͷ
  	
  writeLogFile checkLog, "kill excel and wps that open csv file"
  KillExcelProcess  
  crt.Sleep 200
                                                                                                                
    objFile.Close    
    Set objFSO =Nothing     	 
    For ii_1=0 To UBound(arrDeviceIP)-LBound(arrDeviceIP) '1th#, 
        writeLogFile checkLog, "now ,we check 86x_mem IP-" & arrDeviceIP(ii_1)
	   If crt.Session.Connected Then crt.Session.Disconnect        ' #������ѽ�����������Ͽ����ӡ�
       cmd = "/ssh2 /ACCEPTHOSTKEYS /L root" & " /PASSWORD root" & " /C 3DES " & arrDeviceIP(ii_1)
       numVar=0
       do while numVar < 3  '3������
          err.clear
    	   
    	   if numVar=0 then
    	     crt.Session.ConnectInTab cmd   '��һ����tab������ 
    	   else 
    	     crt.Session.Connect cmd		  '��2,3���ڱ����������� 
    	   end if
    	   'crt.sleep 3000
         If Err.Number <> 0 Then  '��¼������           
             numVar=numVar+1
    	     if numVar > 3 or numVar =3 then 
    	       writeLogFile checkLog, "Exit for 3 times fail!-" & arrDeviceIP(ii_1)
               if g_szSkippedTabs = "" then
                   g_szSkippedTabs = crt.Window.Caption  & vbcrlf '������ΪnIndex
               else
                   g_szSkippedTabs = g_szSkippedTabs & "," & crt.Window.Caption & vbcrlf
               end if		
			      str_1=crt.Window.Caption & ", /,  ssh failed" 
           	      writeResultFile checkResult, str_1 '���浽�ļ�
    	       exit do '������վѲ��
    	     End if
    	     'crt.sleep 3000
          Else   '��¼�ɹ�
    	   numVar =3

           Dim nIndex
           nIndex = 1 '2th#, 
               Set objCurrentTab = crt.GetTab(nIndex)
               objCurrentTab.Activate
               ' Skip tabs that aren't connected
               if objCurrentTab.Session.Connected = True then
       	         
       	         'do sth, end
       	          if ( "NONE" = UCase(cmdUsername) ) then '���û���
                     SendExpect nIndex,"telnet 127.1 2650", "Password:"
	                 SendExpect nIndex, cmdPassword, ">"
	                 SendExpect nIndex,"en", "#"
                  Else  '���û���
                     SendExpect nIndex,"telnet 127.1 2650", "Username:"
	                 SendExpect nIndex, cmdUsername, "Password:"
	                 SendExpect nIndex, cmdPassword, ">"
	                 SendExpect nIndex,"en", "#"
                  End if
       	          szData = CaptureOutputOfCommand(nIndex, "show tne board", "#")
       	          SendExpect nIndex,"exit", "root" 
       	          
       	          'Step2: get board list
                   MyArray = Split(szData, vbcrlf, -1, 1)
                   j=0 '�����������Ͳ�λ
				   writeLogFile checkLog, "read all lines, get NP and scu board slot number--"
                   For ii_3=0 To UBound(MyArray,1)   '3th#, 
       	                s1=instr(1,MyArray(ii_3),"0x380")  'memoryѲ��ͳ��SCU/NP��
                         If s1 > 0 Then
                         	 dim binArray
                         	 
                         	 tmp=""  '�滻�����ո�
                         	 MyArray(ii_3)=Trim(MyArray(ii_3)) 'ȥ����ʼ��β���ո�
                             MyArray(ii_3)=replace(MyArray(ii_3)," ",",")
                             for ii_4 = 1 to len(MyArray(ii_3))  '4th#, 
                              i3=ii_4+1
                              c1=mid(MyArray(ii_3),ii_4,1)  'ȡ��ǰ�ַ�
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
                   j=j-1 '�����1
       	          For ii_5=0 To j   '5th#, �����У���ȡ��λ��Ϣ
       	            writeLogFile checkLog, "read all lines, telnet all NP/SCU--" &  Board(ii_5,0)
					SendExpect nIndex,"telnet 10.26.0." & Board(ii_5,0), "VxWorks login: " 
       	            SendExpect nIndex,"bmu852", "Password: " 
       	            SendExpect nIndex,"aaaabbbb", "->" 
       
       	            szData = CaptureOutputOfCommand(nIndex, "memShow", "->")  '�ж�1��memShow
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
                   tmp=""  '4th#, �滻�����ո�
                   binString=Trim(freeStr) 'ȥ����ʼ��β���ո�
                   binString=replace(binString," ",",")
                   for ii_4 = 1 to len(binString) 
                    i3=ii_4+1
                    c1=mid((binString),ii_4,1)  'ȡ��ǰ�ַ�
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
                    
       	            '��λ��+����+maxBlock+memoryͳ��
       	            Board(ii_5,2)=MaxBlockResult  'maxBlock
                    Board(ii_5,3)=freeBytesNumber 'freeBytes
       	            
       	            SendExpect nIndex,"logout", "root"   '�˳�����

       	           Next  '5th#_Next, 
				   objCurrentTab.Screen.Send "exit" & vbcr    '�˳��Ự
       
       	          writeLogFile checkLog, "save check result,NeName BoardName slot maxBlock freeMemory"
       	          str_1=""
           	      For ii_6=0 To j  '6th#, 
                    str_1=crt.Window.Caption & "," &  Board(ii_6,1) & "," &  Board(ii_6,0) &  _
                          "," &  Board(ii_6,2) & "," &  Board(ii_6,3) 'Chr(13)
                    writeResultFile checkResult, str_1 '���浽�ļ�
				    writeLogFile checkLog, str_1 '���浽��־
				  Next   '6th#_Next, 
       	        'do sth, end
              End if
           '2th#_Next, 
		 end if
		Loop
    Next    '1th#_Next, 
    g_objTab.Activate
    
    if g_szSkippedTabs <> "" Then   'Ѳ�����
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