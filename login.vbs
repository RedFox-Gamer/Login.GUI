a = inputbox("Password:")
    if a = "test" then 'Define your password here
        msgbox "Successful login" 
        CreateObject("WScript.Shell").Run("https://www.instagram.com/_adam.stark_") 'Edit the link you want to open here
    else
        msgbox "Wrong Password"

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run("login.vbs") 'Edit the .vbs's name here if it's different for you
end if