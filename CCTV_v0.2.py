import os, openpyxl, time, pafy, validators, datetime, moviepy.editor, threading

### If you added vlc folder to your PATH just ncomment this
try:
    os.add_dll_directory(openpyxl.load_workbook('PlayList.xlsx', data_only=True)['config'].cell(1, 2).value)
except:
    print('no config sheet found or no path to vlc found in excel cell (1,2)')
    exit()
###

import vlc

YOUTUBE_PREFIX = 'https://www.youtube.com/'
PATH_TO_EXCEL_FILE = 'PlayList.xlsx'
VLC_INSTANCE = vlc.Instance('--avcodec-hw=none --video-on-top --no-directx-hw-yuv')
MEDIA_PLAYER = VLC_INSTANCE.media_player_new()
DEFAULT_VLC_PATH = r'C:\Program Files (x86)\VideoLAN\VLC'
DEFAULT_STATIC_PICTURE = 'https://i.ytimg.com/vi/2vv9IxwFVJY/maxresdefault.jpg'
CONFIG_SHEET_NAME = 'config'
WAITING_TIME_FOR_PLAYER_OPENING = 1.5
DATE_TIME_COLUMN = 1
TITLE_COLUMN = 2
LOCATION_COLUMN = 3
CONFIG_TITLE_COLUMN = 1
CONFIG_VALUE_COLUMN = 2
FIVE_DAYS = 5


class Configuration:
    def __init__(self, vlc_path, static_picture):
        self.static_picture = static_picture
        self.vlc_path = vlc_path


class Media:
    def __init__(self, date_time, title, location, duration, type):
        self.date_time = date_time
        self.title = title
        self.location = location
        self.duration = duration
        self.type = type

    def print(self):
        print((self.date_time, self.title, self.location, self.type))


# Media object of the static_picture
static_picture_media = DEFAULT_STATIC_PICTURE

# In memory representation of complete Excel file
workbook_data = ""

# List of media files to play, the playlist que
playlist_queue = []

# Configuration for the program
config = Configuration(DEFAULT_VLC_PATH, DEFAULT_STATIC_PICTURE)

# Media types dictionary
media_type = dict()
media_type['youtube_video'] = 'YouTube Video'
media_type['local_file'] = 'Local file'
media_type['not_supported'] = 'Not Supported'

# Errors dictionary - list of all possible errors
errors = dict()
errors['no_excel_found'] = "The configuration Excel document {} couldn't be opened or does not exists"
errors['no_valid_youtube'] = "The location {} wasn't found on your computer and is not a valid youtube URL. Was supposed to bo played at {}"
errors['no_valid_datetime'] = "{} is not in a valid datetime format. Found on line number {}"
errors['config_is_invalid'] = "Configuration data in Excel file is missing or wrong"
errors['past_time'] = "{} was supposed to be played at {} which is in the past"


# Checks if the given parameter is in type "datetime.datetime"
def is_time(time):
    try:
        return type(time) == datetime.datetime
    except:
        return False


# Load Excel file to memory
def load_excel_to_memory():
    global workbook_data
    try:
        workbook_data = openpyxl.load_workbook(PATH_TO_EXCEL_FILE, data_only=True)
    except:
        print(errors['no_excel_found'])
        return False
    return True


# Generate configuration from Excel
def generate_configuration_from_excel():
    global config
    try:
        config_workbook = workbook_data[CONFIG_SHEET_NAME]
        for line in range(1, config_workbook.max_row + 1):
            config_title = config_workbook.cell(line, CONFIG_TITLE_COLUMN).value
            config_value = config_workbook.cell(line, CONFIG_VALUE_COLUMN).value
            if config_title == 'VLC-PATH' and os.path.exists(config_value):
                config.vlc_path = config_value
            elif config_title == 'STATIC-PICTRE-PATH' and os.path.exists(config_value):
                config.static_picture = config_value
        return True
    except:
        print(errors['config_is_invalid'])
        return False


# Generates a playlist queue from the excel file
def generate_playlist_queue():
    global playlist_queue
    sheet = workbook_data['playlist']
    present_time = datetime.datetime.now()
    for row in range(2, sheet.max_row + 1):
        date_time = sheet.cell(row, DATE_TIME_COLUMN).value
        title = sheet.cell(row, TITLE_COLUMN).value
        location = sheet.cell(row, LOCATION_COLUMN).value
        type = get_media_type(location)
        duration = get_duration(type, location)
        if not is_time(date_time):
            print(errors['no_valid_datetime'].format(date_time,row))
        elif type != media_type['not_supported'] and date_time > present_time:
            media_object = Media(date_time, title, location, duration, type)
            playlist_queue.append(media_object)
        elif date_time < present_time:
            print(errors['past_time'].format(location, date_time))
        else:
            print(errors['no_valid_youtube'].format(location, date_time))
    playlist_queue.sort(key=lambda media_object: media_object.date_time, reverse=True)
    return True


# Gets the type of the video(local/youtube) and the video location and returns the video duration
def get_duration(type, location):
    if type == media_type['local_file']:
        try:
            return moviepy.editor.VideoFileClip(location).duration
        except:
            return 0
    elif type == media_type['youtube_video']:
        return pafy.new(location).length
    return 0


# Return media type by location string provided
def get_media_type(location):
    try:
        if not location:
            return media_type['not_supported']
        if validators.url(location) and location.find(YOUTUBE_PREFIX) == 0:
            return media_type['youtube_video']
        if os.path.exists(location):
            return media_type['local_file']
        return media_type['not_supported']
    except:
        return media_type['not_supported']


# Gets a media object and plays it
# Currently commented out - code for skipping to the next media if current media is of "not supported" type
# This check is done during the queue generation
def play_media(media_object):
    object_media_type = media_object.type
    media_location = media_object.location
    media_duration = media_object.duration

    # if object_media_type == media_type['not_supported']:
    #     print(errors['no_valid_youtube'], media_location)
    #     return

    if object_media_type == media_type['youtube_video']:
        try:
            video = pafy.new(media_location)
            best = video.getbest()
            media_location = best.url
        except Exception as e:
            print("exception while trying to youtube the video {} \n".format(media_location), e)

    try:
        MEDIA_PLAYER.set_fullscreen(True)
    except Exception as e:
        print("exception while trying to full screen the video {} \n".format(media_location),e)
    try:
        media = VLC_INSTANCE.media_new(media_location)
    except Exception as e:
        print("exception while trying to media new the video {} \n".format(media_location),e)
    try:
        MEDIA_PLAYER.set_media(media)
    except Exception as e:
        print("exception while trying to set media the video {} \n".format(media_location),e)
    try:
        MEDIA_PLAYER.play()
    except Exception as e:
        print("exception while trying to play the video {} \n".format(media_location),e)
    time.sleep(WAITING_TIME_FOR_PLAYER_OPENING)
    print(media_object.date_time,"playing:", media_object.location)
    if not media_location == config.static_picture:
        threading.Timer(media_duration, end_of_media).start()


# Checks if the player currently at "end of media" state and displays the static picture if so.
def end_of_media():
    if MEDIA_PLAYER.get_state() == vlc.State.Ended:
        play_media(static_picture_media)


#Plays the next media in the queue
def play_next_media_in_queue(next_media):
    global playlist_queue
    if not playlist_queue:
        return
    play_media(next_media)
    when_to_play_next = datetime.timedelta(0).total_seconds()
    while when_to_play_next <= datetime.timedelta(0).total_seconds():
        try:
            next_media = playlist_queue.pop()
            if is_time(next_media.date_time):   # TODO check with sleep const (min)
                when_to_play_next = (next_media.date_time - datetime.datetime.now()).total_seconds()
        except:
            return
    threading.Timer(when_to_play_next, play_next_media_in_queue, [next_media]).start()


# Generates static picture dummy media
def generate_static_picture_media_object():
    global static_picture_media
    date_time = datetime.datetime(1,2,3,4,5,6,7)
    title = "static_picture"
    location = config.static_picture
    type = get_media_type(location)
    static_picture_media = Media(date_time, title, location, 0, type)


# loads the excel file, generates configuration and playlist queue from it and generates static-picture media
def init():
    if not load_excel_to_memory():
        exit_handler()
    if not generate_configuration_from_excel():
        exit_handler()
    if not generate_playlist_queue():
        exit_handler()
    generate_static_picture_media_object()


# Exit handler waits for entering any key before the program ends
def exit_handler():
    input("please enter any key to exit")
    exit()


def main():
    global static_picture_media
    init()
    play_next_media_in_queue(static_picture_media)


if __name__ == "__main__":
    main()
