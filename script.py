import win32com.client 
import time 

#Shell Connection 
shell = win32com.client.Dispatch("WScript.Shell")
shell.Run("notepad")
time.sleep(1)
shell.AppActivate("Notepad")


#Message
msg="""Hello World


elena@****.com
12:50 PM : 10 minutes ago
to me

To create is to recombine’ said the molecular biologist Francois Jacob. Imagine if the man who invented the railway and the man who invented the locomotive could never meet or speak to each other, even through third parties. Paper and the printing press, the internet and the mobile phone, coal and turbines, copper and tin, the wheel and steel, software and hardware. I shall argue that there was a point in human pre-history when big-brained, cultural, learning people for the first time began to exchange things with each other, and that once they started doing so, culture suddenly became cumulative, and the great headlong experiment of human economic ‘progress’ began. Exchange is to cultural evolution as sex is to biological evolution.

From The Rational Optimist
by Matt Ridley
""" 
#For Loop for sending characters in sequence in delay

delay=0.04 
for i in msg:
     time.sleep(delay)
     shell.Sendkeys(i,0)

