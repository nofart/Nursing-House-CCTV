import os

os.add_dll_directory(r'C:\Program Files (x86)\VideoLAN\VLC')
import vlc, time, pafy, validators, schedule, openpyxl, datetime


def getCol():
    return datetime.datetime.today().weekday() + 10 % 7  # 1 will be sunday

def getDateFromExcel(wb,col):
    date = wb.active.cell(1,col).value
    return date


def checkIfTime(cell):
    try:
        return type(cell.value) == datetime.time
    except:
        False

def createScheduleList(excel):
    global schedule_list,times
    wb = openpyxl.load_workbook(excel, data_only=True)
    for sheet in wb:
        for row in range(2,wb.active.max_row):
            time = wb.active.cell(row,1).value
            if type(time) == datetime.time:
                times.append(time)
                for col in range(2, wb.active.max_column):
                    value = getLocation(wb, row + 1, col)
                    if value:
                        date = getDateFromExcel(wb, col).date()
                        schedule_list[(date, time)]=value


def runTask(t):
    global schedule_list
    today = datetime.datetime.today().date()
    key = (today,t)
    location = schedule_list.get(key)
    openVideo(location)



def scheduleTasks():
    global times
    for t in times:
        schedule.every().day.at(t.strftime("%H:%M")).do(runTask, t)


def getLocation(wb,row, col):
    return wb.active.cell(row, col).value

excel = '3.xlsm'
static_picture = 'lev.jpg'  # path to the default picture to share on screen
schedule_list = {}
times=[]

def play(location):
    Media = Instance.media_new(location)
    Media.get_mrl()
    player.set_media(Media)
    player.play()

def openVideo(location):
    global player, Instance, static_picture
    print(location)
    if not location:
        return
    if validators.url(location):
        video = pafy.new(location)
        best = video.getbest()
        duration = pafy.new(location).length
        location = best.url
    if os.path.exists(location):
        time.sleep(1.5)
        duration = player.get_length() / 1000
    time.sleep(1.5)
    play(location)
    time.sleep(duration)
    play(static_picture)


Instance = vlc.Instance(['--video-on-top', '--keyboard-events', '--mouse-events'])
player = Instance.media_player_new()
player.set_fullscreen(True)
play(static_picture)

createScheduleList(excel)
print(schedule_list)
print(times)
scheduleTasks()
schedule.every().day.at("00:00").do(scheduleTasks)

while True:
    schedule.run_pending()
    time.sleep(1)