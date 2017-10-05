import win32api, win32con, threading
from tkinter import Tk
from time import sleep
from tkinter.filedialog import askopenfilename, asksaveasfilename
from os import system, remove


class Presser_Thread(threading.Thread):
	def __init__(self, instruction_set):
		threading.Thread.__init__(self)
		self.instruction_set = instruction_set
		self.quit = False

	def run(self):
		print("Starting presser.")
		sleep(1)
		while not self.quit:
			for key, time in self.instruction_set:
				win32api.keybd_event(key, 0, win32con.KEYEVENTF_EXTENDEDKEY, 0)
				sleep(time/1000)
				if self.quit:
					break
		print("Presser stopped.")


def instruction_set_from_file(file_path):
	try:
		with open(file_path, "r") as f:
			instruction_set = []
			for line in f.readlines():
				line = line.strip()
				#if not comment line or blank
				if line and not line[0] == "#":
					#if there is a space, attempt to read a specified millisecond amount
					if " " in line:
						parts = line.split()
						line = parts[0]
						wait = int(parts[1])
					else:
						wait = 1000

					if line[0] != "-":
						code = ord(line[0])
						#capital letter or digit needs no editing
						if 65 <= code <= 90 or 48 <= code <= 57:
							pass
						#lowercase letter needs to be decremented by 32
						elif 97 <= code <= 122:
							code -= 32
						else:
							raise ValueError("Invalid code")
					else:
						code = int(line)
						if not 0 <= code <= 254:
							raise ValueError("Invalid code")

					instruction_set.append((code, wait))

		if instruction_set:
			for line in instruction_set:
				print(line)
			return instruction_set
		else:
			return None
	except (FileNotFoundError, ValueError):
		return None


def main():
	Tk().withdraw()
	instruction_set = None
	system("color 71")

	try:
		with open("autoload.ini", "r") as f:
			current_file = f.readline().strip()
		instruction_set = instruction_set_from_file(current_file)
		if not instruction_set:
			print("Invalid autoload.ini file, deleting it.")
			remove("autoload.ini")
			current_file = None
	except FileNotFoundError:
		current_file = None

	inp = ""
	while inp != "q":
		inp = input("$ ")

		if inp == "help":
			print("Commands:\n" + 
				"open  - choose a file to read a key press sequence from\n" +
				"save  - save the currently chosen file as the default file to read on startup\n" +
				"start - start executing the key press sequence that is loaded\n" +
				"q     - exit the program")

		elif inp == "open":
			current_file = askopenfilename()
			instruction_set = instruction_set_from_file(current_file)
			if not instruction_set:
				print("No file was chosen or the file could not be read.")
				current_file = None

		elif inp == "save":
			if current_file:
				with open("autoload.ini", "w") as f:
					f.write(current_file)
				print("Saved auto load file as %s." % current_file)
			else:
				print("Cannot save because no valid key press sequence is loaded.")

		elif inp == "start":
			if instruction_set:
				presser_thread = Presser_Thread(instruction_set)
				presser_thread.start()
				input("Press enter to stop key pressing.")
				presser_thread.quit = True
				while presser_thread.isAlive():
					sleep(0.05)
			else:
				print("Cannot start because no valid key press sequence is loaded.")

		elif inp == "q":
			continue

		else:
			print("'%s' is not a recognized command. Enter 'help' to list valid commands." % inp)

	Tk().destroy()

main()