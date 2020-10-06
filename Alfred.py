import datetime
import win32com.client
import os
import time

speaker = win32com.client.Dispatch("SAPI.SpVoice")
txtfile = "test.txt"
days = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
bot = True

def saludateInit():
    dateNow = datetime.datetime.now()
    day = dateNow.strftime('%a')
    #print ("Formato dd/mm/yyyy =  %s/%s/%s" % (dateNow.day, dateNow.month, dateNow.year))
    #print ("Formato hh:mm = %s:%s" % (dateNow.hour, dateNow.minute))
    saludate = compareHour(dateNow.hour)
    dayNow = weekDay(day, days)
    message = ("Hola Fran "+ saludate + " espero que estes muy bien, estare al pendiente de tus tareas del día "
    + dayNow + "," + str(dateNow.day) + ". Avisame si necesitas algo.")
    return message
    
def compareHour(hour):
    if(hour < 12):
        return "buenos días"
    elif(hour >= 12 and hour < 18):
        return "buenas tardes"
    elif(hour >= 18):
        return "buenas noches"

def weekDay(weekday, days):
    if(days[0] == weekday):
        day = "Lunes"
    elif(days[1] == weekday):
        day = "Martes"
    elif(days[2] == weekday):
        day = "Miércoles"
    elif(days[3] == weekday):
        day = "Jueves"
    elif(days[4] == weekday):
        day = "Viernes"
    elif(days[5] == weekday):
        day = "Sábado"
    elif(days[6] == weekday):
        day = "Domingo"
    return day

def writeTxt(txtfile, text):
    file = open(txtfile, "a")
    file.write(text)
    file.close()
    
def readTxt(txtfile):
    file = open(txtfile, "r")
    for line in file.readlines():
        dateTime = datetime.datetime.now()
        dateNow = str(dateTime.day) + "/" + str(dateTime.month) + "/" + str(dateTime.year)
        hourNow = str(dateTime.hour)
        minuteNow = str(dateTime.minute)
        text = line.split("] ")
        cutDateHour = text[0].split("-")
        cutDate = cutDateHour[0].split("[")
        date = cutDate[1]
        cutHourMinute = cutDateHour[1].split(":")
        hour = cutHourMinute[0]
        minute = cutHourMinute[1]
        activity = text[1]
        if(dateNow == date):
            time = int(hour) - int(hourNow)
            if(time <= 1):
                clock = int(minute) - int(minuteNow)
                if(clock <= 30):
                    message = "En menos de una hora será su actividad: " + activity
                    speaker.Speak(message)
                    print(line)
                    print("---------------------------------------------------------------------------------")
    file.close()
    
def clearScreen():
    os.system("clear")

def createTasks():
    date = input("Ingrese la fecha de la actividad (dd/mm/yyyy): ")
    hour = input("Ingrese la hora de la actividad (hh:mm): ")
    task = input("Indique la actividad: ")
    text = "[" + date + "-" + hour + "]" + " " + task + "\n"
    writeTxt(txtfile, text)
    
def alertTasks():
    print("Cada media hora estoy pendiente de si tienes alguna actividad...")
    while(True):
        readTxt(txtfile)
        time.sleep(900)

message = saludateInit()
speaker.Speak(message)
while(bot):
    clearScreen()
    print("-----------------------------------------------------------")
    print("|   Soy Frosty, tu servidora personal. Dime que quieres:  |")
    print("-----------------------------------------------------------")
    print("1. Ingresar actividad")
    print("2. Mantente pendiente de mis actividades")
    print("9. Salir")
    print("-----------------------------------------------------------")
    response = int(input("Indicame el número de la opción que deseas: "))
    if(response == 1):
        clearScreen()
        createTasks()
    elif(response == 2):
        clearScreen()
        alertTasks()
    else:
        clearScreen()
        bot = False