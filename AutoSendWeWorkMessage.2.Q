[General]
SyntaxVersion=2
BeginHotkey=121
BeginHotkeyMod=0
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=cf18498f-d2a4-47e9-b7d6-a803a227ae27
Description=Auto Send QYWX Message.2
Enable=0
AutoRun=0
[Repeat]
Type=0
Number=1
[SetupUI]
Type=2
QUI=
[Relative]
SetupOCXFile=
[Comment]

[Script]
Dim screen_height,screen_width
screen_width = Plugin.sys.GetScRX()
screen_height = Plugin.sys.GetScRY() 
TracePrint "screen width:"&screen_width&",height:"&screen_height

Dim excel_cfg_file,process_model
Dim read_line_num, start_read_line, send_msg_num
Dim group_users, group_name, send_msg
send_msg = Array("", "", "", "")
//��s
excel_cfg_file = "D:\Test002.xlsx"

send_msg_num = 3
start_read_line = 2
end_read_line = 56

process_model = 1

read_line_num = end_read_line - start_read_line + 1

If send_msg_num > 3 And send_msg_num <= 0 Then  
	TracePrint "send_msg_num(must <3):" & send_msg_num
	MessageBox "һ�θ�һ��Ⱥ���͵���Ϣ�����������������ܹ���1<number<3"
	ExitScript
End If

If read_line_num <= 0 Or start_read_line <=0 Or end_read_line <=0 Then  
	TracePrint "read_line_num(must >0):" & read_line_num
	MessageBox "һ�θ�һ��Ⱥ���͵���Ϣ�����������������ܹ���1<number<3"
	ExitScript
End If

Call ShowWeWorkWnd()

If process_model = 0 Then 
	Call CreateGroupAndSendMsg()
ElseIf	 process_model = 1 Then
	Call FindGroupAndSendMsg()
ElseIf	 process_model = 2 Then
	Call FindUserAndSendMsg()
Else 
	TracePrint "process_model error must equal 0,1,2" & process_model	 
End If
ExitScript

Function ShowWeWorkWnd()
	Dim hwnd_qw_main,hwnd_fg
	hwnd_qw_main = Plugin.Window.Find("WeWorkWindow", 0)
	//Call Plugin.Window.Top(hwnd_qw_main, 0)
	//Delay 1000
	Call Plugin.Window.Active(hwnd_qw_main)
	Delay 2500
	hwnd_fg = Plugin.Window.Foreground()
	If hwnd_fg = hwnd_qw_main Then
		TracePrint ("Active QY Weixin OK.")
	Else
		TracePrint ("Active QY Weixin  No OK.")
		MessageBox "�޷�������ҵ΢�Ŵ���"
		ExitScript
	End If
	
	Dim rect, rect_array
	rect = Plugin.Window.GetClientRect(hwnd_qw_main)
	rect_array = Split(rect, "|")
	wleft = Clng(rect_array(0)): wtop = Clng(rect_array(1))   
	wright = Clng(rect_array(2)) : wbottom = Clng(rect_array(3))
	TracePrint "Rect  " & wleft & "," & wtop & "," & wright & "," & wbottom
End Function

//����Ⱥ��Ȼ������Ϣ
Function CreateGroupAndSendMsg()
	Dim read_num
	read_num = 0
	Call Plugin.Office.OpenXls(excel_cfg_file)
	While read_num < read_line_num
    	
		Call ReadConfigFromExcel(start_read_line + read_num)
		
    	Call OpenCreateGroupChatWnd()
    	Call CreateGroupChat(group_users)
    	Call RenameGroupChatMainWnd(group_name)
    	Call EnterFindWnd()
    	Call FindGroupChatWnd(group_name)
		Call BatchSendWeWorkMessage()
    	
    	read_num = read_num + 1
	Wend
	Call Plugin.Office.CloseXls()
End Function

Function FindGroupAndSendMsg()
	Dim read_num
	read_num = 0
	Call Plugin.Office.OpenXls(excel_cfg_file)
	While read_num < read_line_num
    	
		Call ReadConfigFromExcel(start_read_line + read_num)
    	Call EnterFindWnd()
    	Call FindGroupChatWnd(group_name)
		Call BatchSendWeWorkMessage()
    	read_num = read_num + 1
	Wend
	Call Plugin.Office.CloseXls()
End Function

Function FindUserAndSendMsg()
	
End Function

Function ReadConfigFromExcel(line_num)
	group_users = Plugin.Office.ReadXls(1, line_num, 1)
    group_name = Plugin.Office.ReadXls(1, line_num, 2)
    If group_users = "" Or group_name = "" Then 
    	TracePrint ("EXCEL have not config.")
		MessageBox "EXCEL ��û����Ӧ������"
		ExitScript
	End If
	If send_msg_num = 1 Then 
		send_msg(0) = Plugin.Office.ReadXls(1, line_num, 3 + 0)
	ElseIf send_msg_num = 2 Then 
		send_msg(0) = Plugin.Office.ReadXls(1, line_num + read_num, 3 + 0)
		send_msg(1) = Plugin.Office.ReadXls(1, line_num + read_num, 3 + 1)
	// = 3
	Else 
		send_msg(0) = Plugin.Office.ReadXls(1, line_num, 3 + 0)
		send_msg(1) = Plugin.Office.ReadXls(1, line_num, 3 + 1)
		send_msg(2) = Plugin.Office.ReadXls(1, line_num, 3 + 2)
		//TracePrint send_msg(0)
		//TracePrint send_msg(1)
		//TracePrint send_msg(2)
	End If
End Function

//Ctrl + N �򿪴���Ⱥ�ĵĴ���
Function OpenCreateGroupChatWnd()
	KeyDown 17, 1
    Delay 100
    KeyPress 78, 1
    Delay 100
    KeyUp 17, 1
    Delay 2000
End Function

//Ctrl + Enter ������Ϣ
Function BatchSendWeWorkMessage()
    If send_msg_num = 1 Then 
		Call SendWeWorkMessage(send_msg(0))
	ElseIf send_msg_num = 2 Then 
		Call SendWeWorkMessage(send_msg(0))
		Call SendWeWorkMessage(send_msg(1))
	// = 3
    Else 
		Call SendWeWorkMessage(send_msg(0))
		Call SendWeWorkMessage(send_msg(1))
		Call SendWeWorkMessage(send_msg(2))
	End If
End Function

//Ctrl + Enter ������Ϣ
Function SendWeWorkMessage(one_msg)
    SayString one_msg
    Delay 1500
    //����ط�ѡ���Ctrl + Enter ���ͣ������ط�����
	KeyDown 17, 1
    Delay 30
    KeyPress "Enter", 1
    Delay 30
    KeyUp 17, 1
    Delay 1500
End Function

//����Ⱥ��
Function CreateGroupChat(group_users)
	Plugin.Sys.SetCLB(group_users)
 	KeyDown 17, 1
 	 Delay 100
	KeyPress 86, 1
	 Delay 100
	KeyUp 17, 1
	Delay 3500
    KeyPress "Enter", 2
    Delay 1500
End Function

//��Ⱥ�����ڸ���
Function RenameGroupChatMainWnd(group_name)
	TracePrint "RenameGroupChatMainWnd"
	//���ݴ���"+"��λ
	Dim xy_array,xy,x,y
	xy = Plugin.Color.FindMultiColorEx(wleft,wtop,wright,wtop+90,"000000","0|0|000000,1|0|000000,0|1|000000,1|1|000000,0|2|000000,1|2|000000,0|3|000000,1|3|000000",1,0)
	xy_array = Split(xy, "|")
	x = CInt(xy_array(0)) : y = CInt(xy_array(1))
	TracePrint "Fild xy "&xy&" "& x &"," &y
	//���û���ҵ����˳�
	If x < 0 Or y < 0 Then 
		TracePrint "Not Fild Group Chat Name Windows "
		ExitScript
	End If
	//����Ⱥ���ƴ��ڣ�
    MoveTo x, y
    Delay 200
    LeftClick 1
    SayString group_name
	Delay 200
	MoveR 150, 0
    LeftClick 1
    Delay 200
End Function

//��Ⱥ�����ڸ���
Function RenameGroupChatWnd(group_name)
    hwnd_qw_c = Plugin.Window.Find("WwStandaloneConversationWnd", 0)
    Call Plugin.Bkgnd.LeftClick(hwnd_qw_c, 30, 40)
	Delay 500
    SayString group_name
   Delay 1500
End Function

//Ctrl+O 
Function OpenGroupChatWnd()
	TracePrint "Ctrl+O OpenGroupChatWnd"
	KeyDown 17, 1
    Delay 30
    KeyPress 79, 1
    Delay 30
    KeyUp 17, 1
    Delay 1000
End Function

//Esc
Function CloseGroupChatWnd()
	KeyPress "Esc", 1
    TracePrint "Esc CloseGroupChatWnd"
    Delay 1000
End Function

Function  EnterFindWnd()
    KeyPress 18,1
    Delay 30
    KeyPress 18,1
    Delay 30
    TracePrint "Alt + Alt EnterFindWnd"
    Delay 500
End Function

Function FindGroupChatWnd(group_name)
	TracePrint "FindGroupChatWnd "
	//ȫѡ���е����ݣ���ɾ��ԭ�������ݣ�
    KeyDown 17, 1
    Delay 30
	KeyPress 65, 1
	Delay 30
	KeyUp 17, 1
	Delay 500
	SayString group_name
	Delay 1500
    //�����Ϊ�Ǹ��Ŵ󾵷ǳ�����������û���ҵ�
    Dim xy_array,xy,x,y
	xy =Plugin.Color.FindMultiColorEx(wleft,wtop,wleft+300,wtop+160,"56B059","0|0|56B059,1|0|56B059,2|0|56B059,0|1|56B059,1|1|56B059,2|1|56B059,0|2|56B059,1|2|56B059,2|2|56B059,3|2|56B059,3|2|56B059,3|2|56B059",0.9,0)
	xy_array = Split(xy, "|")
	x = CInt(xy_array(0)) : y = CInt(xy_array(1))
	TracePrint "Fild xy " & xy &","& x &"," &y
	//����ҵ���ɫ�Ŵ��ӣ���ʾ����������ʵû���ҵ������е��ţ����˳�
	If x > 0 Or y > 0 Then 
		TracePrint "FindGroupChatWnd Not find group :["&group_name&"] to send message"
		ExitScript
	End If
	KeyPress "Enter", 1
	Delay 500	
End Function

// �ű�����ʱ��ʾ
Sub OnScriptExit()
    KeyUp 17, 1
    Call Plugin.Office.CloseXls()
    MessageBox "�ű��Ѿ�ֹͣ��"
End Sub