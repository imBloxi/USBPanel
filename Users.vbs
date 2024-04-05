dim users_list
user_number = 1

Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set login_list = objWMIService.ExecQuery("Select * from Win32_NetworkLoginProfile")

For Each login in login_list 
    users_list = users_list & vbCr & user_number & " " & login.Name 
    user_number = user_number + 1
Next

WScript.echo users_list
WScript.quit
