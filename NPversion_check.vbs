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
cmdUsername="none"  'telnet RCU ���û���
cmdPassword="fiberhome" 'telnet RCU������

Dim arrDeviceIP(), arrCMD(), ipNumber
countOk=0 '����վ����
countBad=0 '�쳣վ����
ipNumber = 0 '��¼IPվ��ĸ���

Dim Board() '���ȶ���һ��һά��̬����
'��λ��+����+R1X+CRCͳ��
ReDim Board(50,3) '���¶���Ϊ��ά����,51��4�С�

Sub Main
    On Error Resume Next
	     If crt.Dialog.MessageBox(_
        "check NP board after 3s?" & vbcrlf & vbcrlf , _
        "check NP board", _
        vbyesno) <> vbyes then exit Sub

    '��һ������ȡpingSetings�ļ�����
    Dim currentPath, objFSO, MyArray, binArray, R86xYesNo
       Set objFSO = CreateObject("Scripting.FileSystemObject")
      currentPath = createobject("Scripting.FileSystemObject").GetFolder(".").Path + "\"      '��ȡ��ǰ��·��
    '--------------------------------------------------------------------
      szData =  now '��¼���β���ʱ��
      szData = Replace(szData, "/", "-")  '�л���
      szData = Replace(szData, " ", "_")  '�»���
      szData = Replace(szData, ":", "-")  '�л���

      fileDate= currentPath&"NPversion_checkLog" & ".log"
      writeLogFile fileDate, "check starts, NPversion_checkLog-" & fileDate
      writeLogFile fileDate, "Option Time:" & szData

	Set objFile = objFSO.OpenTextFile(currentPath&"IpSeting.txt", 1)
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
      filePath=resultPath&"NPversion_checkResult" & szData & ".csv"
      writeLogFile fileDate, "resultPath exists! --"& resultPath
  else
       filePath=currentPath&"NPversion_checkResult" & szData & ".csv"
       writeLogFile fileDate, "resultPath is none or resultPath donot exist, value is--"& resultPath
  end if

  if (objFSO.fileexists(filePath)) then   '�ж�result.csv�ļ��Ƿ����
  	  objFSO.deletefile(filePath)         'ɾ��result.csv�ļ�
  End if
  writeLogFile fileDate, "NPversion_check result --" & filePath
  str_1="NeIP, boardName, slot, NpVersion, checkTime--" & now 'д�뱨ͷ
  writeResultFile filePath, str_1 'д�뱨ͷ

  writeLogFile fileDate, "kill excel and csv program for reading csv file"
  KillExcelProcess
  crt.Sleep 200

    objFile.Close
    Set objFSO =Nothing
    For ii_1=0 To UBound(arrDeviceIP)-LBound(arrDeviceIP) '1th#,
        writeLogFile fileDate, "NPversion_check NE IP-" & arrDeviceIP(ii_1)
	   If crt.Session.Connected Then crt.Session.Disconnect        ' #������ѽ�����������Ͽ����ӡ�
       cmd = "/ssh2 /L root" & " /PASSWORD root" & " /C 3DES " & arrDeviceIP(ii_1)
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
    	       writeLogFile fileDate, "Exit for 3 times fail!-" & arrDeviceIP(ii_1)
               if g_szSkippedTabs = "" then
                   g_szSkippedTabs = crt.Window.Caption  & vbcrlf '������ΪnIndex
               else
                   g_szSkippedTabs = g_szSkippedTabs & "," & crt.Window.Caption & vbcrlf
               end if
			      str_1=crt.Window.Caption & ", /,  because of not connected�� NE cannot be checked"
           	      writeResultFile filePath, str_1 '���浽�ļ�
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
				 R86xYesNo="none" 'yes-R86x, no-R845
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
				  writeLogFile fileDate, "show tne board --" & szData
				  'get R86x or R845?
				  scur1=0
				  scuo1=0
				  scup1=0
				  scuq1=0
				  R86xYesNo="none"
				  '86x�豸Ѳ�첻ͳ��SCU�̣�����SCU�����ĵ��̶����ˡ�
				  scur1=instr(1,szData,"SCUR1")
				  scuo1=instr(1,szData,"SCUO")
				  '845�豸Ѳ��ֻͳ��SCU�̡�
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
					 'exit do '������վѲ��
				  end if
       	          SendExpect nIndex,"exit", "root"

       	          'Step2: get board list
                   MyArray = Split(szData, vbcrlf, -1, 1)
                   j=0 '�����������Ͳ�λ
				   writeLogFile fileDate, "read show-tne-board to get slot and boardName"
                   For ii_3=0 To UBound(MyArray,1)   '3th#,
       	                s1=instr(1,MyArray(ii_3),"0x380")
						'86x�豸Ѳ�첻ͳ��SCU�̣�����SCU�����ĵ��̶����ˡ�
						scur1=instr(1,MyArray(ii_3),"SCUR1")
						scuo1=instr(1,MyArray(ii_3),"SCUO") 'SCUO1 or SCUO2
						'845�豸Ѳ��ֻͳ��SCU�̡�
       	                scup1=instr(1,MyArray(ii_3),"SCUP1")
						scuq1=instr(1,MyArray(ii_3),"SCUQ1")
                         
						 scuBoardYesNo=False 'case1: SCUX board
						 scuBoardYesNo=0 < scur1 or 0 < scuo1 or 0 < scup1 or 0 < scuq1
						 
						 If s1 > 0 And "none" <> R86xYesNo and True =scuBoardYesNo Then 'case1: SCUX board
						     writeLogFile fileDate,  "board is R845/R86x SCUSX"
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
						 
						 If s1 > 0 And "yes" = R86xYesNo and False =scuBoardYesNo Then 'case2: R86x Np board
                         	 writeLogFile fileDate,  "board is R86x Np"
							 tmp=""  '�滻�����ո�
                         	 MyArray(ii_3)=Trim(MyArray(ii_3)) 'ȥ����ʼ��β���ո�
                            MyArray(ii_3)=replace(MyArray(ii_3)," ",",")
                            for ii_4 = 1 to len(MyArray(ii_3))  '4.1th#,
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
                            Next  '4.1th#_Next,

                         	   MyArray(ii_3)=tmp

                         	 binArray=Split(MyArray(ii_3), ",", -1, 1)
                         	   Board(j,0)=binArray(0)  'slot
                         	   Board(j,1)=binArray(2)  'board name
                         	   j=j+1 'next
                         End If
						 
                   Next   '3th#_Next,

                   writeLogFile fileDate, "telnet 10.26.0.x"
                   j=j-1 '�����1
       	          For ii_5=0 To j   '5th#, �����У���ȡ��λ��Ϣ
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
       	            else '����,���ԭʼֵ
					    Board(ii_5,2)="error��type is not sure��"
					End If

					SendExpect nIndex,"logout", "root"   '�˳����� 

       	          Next  '5th#_Next,

       	          writeLogFile fileDate, "save check result: NeIP	boardName slot	NpVersion"
       	          str_1=""
           	      For ii_6=0 To j  '6th#,
                    str_1=crt.Window.Caption & "," &  Board(ii_6,1) & "," &  Board(ii_6,0) &  _
                          "," &  Board(ii_6,2)  'Chr(13)
                    writeResultFile filePath, str_1 '���浽�ļ�
				    writeLogFile fileDate, str_1 '���浽��־
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
            "because of not connected, NE cannot be checked��" & _
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
    	 'writeFile filePath, str_1
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
'CRC�������¼
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
