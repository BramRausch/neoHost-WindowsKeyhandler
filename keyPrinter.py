# Copyright 2014 Bram Rausch -- GPLv2, See license for reuse detials
import win32com.client
import serial

ser = serial.Serial("COM3", 9600) #open the seri al port
keyb = win32com.client.Dispatch("WScript.shell")

while True:
	letter = ser.read() #read the serial input from the configured port
	
	try: #try to convert var letter to ASCII
		letter.decode("ASCII") 
	except UnicodeError: #if there is a error pass
		pass
	else:
		letter = letter.decode("ASCII") #if converting the input letter does work convert it
		letter.replace("\r", "")
		if letter != " " and letter != '\x08' and letter != '\x00': #some exeptions
			keyb.sendKeys(letter, 0) #send keystroke
