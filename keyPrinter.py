import win32com.client
import serial

ser = serial.Serial("COM3", 9600) #open the serial port
keyb = win32com.client.Dispatch("WScript.shell")

while True:
	letter = ser.read() #read the serial input from the configured port
	
	try: #try to convert var letter to ASCII
		letter = letter.decode("ASCII")
	except UnicodeError: #if there is a error pass
		pass
	else:
		letter.replace("\r", "")
		if letter != '\x08' and letter != '\x00':
			if letter == chr(8): 
				keyb.sendKeys({BS}) #send backspace
			else:
				keyb.sendKeys(letter) #send keystroke
