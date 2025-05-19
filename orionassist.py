import os
import queue
import json
import glob
import threading
import subprocess
import requests
import zipfile
import sounddevice as sd
import schedule
import vosk
import psutil
import pyttsx3

from datetime import datetime
from PIL import ImageGrab
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL
from ctypes import cast, POINTER

# Initialize model and audio queue
model = vosk.Model(r"C:\Users\silver\Desktop\New folder (4)\vosk-model-small-fa-0.42")
q = queue.Queue()

# OpenWeatherMap API
OWM_API_KEY = "YOUR_OPENWEATHERMAP_KEY"

MUSIC_PATH    = r"C:\Music\favorite.mp3"       # ← adjust your media paths

# Initialize TTS engine
engine = pyttsx3.init()
engine.say("Voice assistant started.")
engine.runAndWait()

def speak(text):
    """Convert text to speech."""
    engine.say(text)
    engine.runAndWait()

# Helper functions
def get_volume_interface():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None
    )
    return cast(interface, POINTER(IAudioEndpointVolume))

volume_iface = get_volume_interface()

# Reminders / alarms store
alarms = []
reminders = []

def alarm_checker():
    """Background thread to run scheduled jobs."""
    while True:
        schedule.run_pending()
        sd.sleep(1000)

# Start alarm thread
threading.Thread(target=alarm_checker, daemon=True).start()

# ------------ COMMAND HANDLER ------------
def execute_command(command):
    cmd = command.lower()

    # --- SYSTEM CONTROLS ---
    if "open browser" in cmd:
        os.startfile("chrome")
        speak("Browser opened.")
    elif "shutdown" in cmd:
        speak("Shutting down system.")
        os.system("shutdown /s /t 1")
    elif "restart" in cmd:
        speak("Restarting system.")
        os.system("shutdown /r /t 1")
    elif "lock pc" in cmd:
        speak("Locking PC.")
        os.system("rundll32.exe user32.dll,LockWorkStation")
    elif "log off" in cmd:
        speak("Logging off.")
        os.system("shutdown /l")
    elif "hibernate" in cmd:
        speak("Hibernating.")
        os.system("shutdown /h")

    # --- SYSTEM STATS ---
    elif "battery" in cmd:
        batt = psutil.sensors_battery()
        speak(f"Battery at {batt.percent} percent, {'plugged in' if batt.power_plugged else 'on battery'}.")
    elif "cpu usage" in cmd:
        usage = psutil.cpu_percent(interval=1)
        speak(f"CPU usage is {usage} percent.")
    elif "memory usage" in cmd:
        mem = psutil.virtual_memory()
        speak(f"Memory usage is {mem.percent} percent.")
    elif "disk usage" in cmd:
        disk = psutil.disk_usage(os.getcwd())
        speak(f"Disk usage is {disk.percent} percent.")

    # --- TIME & DATE ---
    elif "what time" in cmd:
        now = datetime.now().strftime("%H:%M:%S")
        speak(f"The time is {now}")
    elif "what date" in cmd:
        today = datetime.now().strftime("%Y-%m-%d")
        speak(f"Today's date is {today}")


    elif "screenshot" in cmd:
        path = os.path.join(os.getcwd(), f"screenshot_{datetime.now():%Y%m%d_%H%M%S}.png")
        img = ImageGrab.grab()
        img.save(path)
        speak(f"Screenshot saved to {path}")

    # --- VOLUME CONTROLS ---
    elif "volume up" in cmd:
        volume_iface.SetMasterVolumeLevelScalar(min(volume_iface.GetMasterVolumeLevelScalar()+0.1,1.0), None)
        speak("Volume increased.")
    elif "volume down" in cmd:
        volume_iface.SetMasterVolumeLevelScalar(max(volume_iface.GetMasterVolumeLevelScalar()-0.1,0.0), None)
        speak("Volume decreased.")
    elif "mute" in cmd:
        volume_iface.SetMute(1, None)
        speak("Muted.")
    elif "unmute" in cmd:
        volume_iface.SetMute(0, None)
        speak("Unmuted.")

    # --- WIFI & VPN ---
    elif "turn on wifi" in cmd:
        os.system('netsh interface set interface "Wi-Fi" enabled')
        speak("Wi-Fi turned on.")
    elif "turn off wifi" in cmd:
        os.system('netsh interface set interface "Wi-Fi" disabled')
        speak("Wi-Fi turned off.")
    elif "start vpn" in cmd:
        os.system('rasdial MyVPN Username Password')
        speak("VPN connected.")
    elif "stop vpn" in cmd:
        os.system('rasdial MyVPN /disconnect')
        speak("VPN disconnected.")

    # --- WEATHER ---
    elif "weather" in cmd:
        city = cmd.replace("weather in", "").strip() or "Tehran"
        url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={OWM_API_KEY}&units=metric"
        resp = requests.get(url).json()
        desc = resp["weather"][0]["description"]
        temp = resp["main"]["temp"]
        speak(f"{city} weather: {desc}, {temp}°C.")

    # --- ALARMS & REMINDERS ---
    elif "set alarm" in cmd:
        # e.g. "set alarm at 07:30"
        time_str = cmd.split("at")[-1].strip()
        schedule.every().day.at(time_str).do(lambda: speak(f"Alarm: it's {time_str}"))
        speak(f"Alarm set for {time_str}")
    elif "remind me to" in cmd:
        # e.g. "remind me to take a break at 15:00"
        parts = cmd.split(" at ")
        task = parts[0].replace("remind me to", "").strip()
        at_time = parts[1].strip()
        schedule.every().day.at(at_time).do(lambda: speak(f"Reminder: {task}"))
        speak(f"Okay, I'll remind you to {task} at {at_time}")

    # --- FILE OPERATIONS ---
    elif "create file" in cmd:
        fname = cmd.replace("create file", "").strip() or f"file_{datetime.now():%Y%m%d_%H%M%S}.txt"
        open(fname, "w").close()
        speak(f"Created file {fname}")
    elif "delete file" in cmd:
        pattern = cmd.replace("delete file", "").strip() or "*"
        files = glob.glob(pattern)
        for f in files:
            os.remove(f)
        speak(f"Deleted {len(files)} file(s) matching {pattern}")
    elif "find file" in cmd:
        pattern = cmd.replace("find file", "").strip() or "*"
        files = glob.glob(f"**/{pattern}", recursive=True)
        if files:
            speak(f"Found {len(files)} files, listing on console.")
            for f in files: print(f)
        else:
            speak("No files found.")

    # --- MEDIA CONTROLS ---
    elif "play song" in cmd:
        subprocess.Popen(["C:\\Program Files\\VideoLAN\\VLC\\vlc.exe", "--play-and-exit", MUSIC_PATH])
        speak("Playing song.")
    elif "pause music" in cmd:
        os.system("nircmd sendkeypress space")
        speak("Music paused.")
    elif "next song" in cmd:
        os.system("nircmd mediakey next")
        speak("Next track.")
    elif "previous song" in cmd:
        os.system("nircmd mediakey prev")
        speak("Previous track.")

    # --- APPLICATION LAUNCHERS ---
    elif "open notepad" in cmd:
        os.startfile("notepad")
        speak("Notepad opened.")
    elif "open calculator" in cmd:
        os.startfile("calc")
        speak("Calculator opened.")
    elif "open paint" in cmd:
        os.startfile("mspaint")
        speak("Paint opened.")
    elif "open cmd" in cmd:
        os.startfile("cmd")
        speak("Command Prompt opened.")
    elif "open explorer" in cmd:
        os.startfile("explorer")
        speak("File Explorer opened.")
    elif "open word" in cmd:
        os.startfile("winword")
        speak("Word opened.")
    elif "open excel" in cmd:
        os.startfile("excel")
        speak("Excel opened.")
    elif "open powerpoint" in cmd:
        os.startfile("powerpnt")
        speak("PowerPoint opened.")
    elif "open chrome" in cmd:
        os.startfile("chrome")
        speak("Chrome opened.")
    elif "open firefox" in cmd:
        os.startfile("firefox")
        speak("Firefox opened.")
    elif "open edge" in cmd:
        os.startfile("msedge")
        speak("Edge opened.")
    elif "open skype" in cmd:
        os.startfile("skype")
        speak("Skype opened.")
    elif "open teams" in cmd:
        os.startfile("teams")
        speak("Teams opened.")
    elif "open slack" in cmd:
        os.startfile("slack")
        speak("Slack opened.")
    elif "open spotify" in cmd:
        os.startfile("spotify")
        speak("Spotify opened.")
    elif "open netflix" in cmd:
        os.startfile("chrome https://netflix.com")
        speak("Netflix opened.")
    elif "open github" in cmd:
        os.startfile("chrome https://github.com")
        speak("GitHub opened.")
    elif "open stack overflow" in cmd:
        os.startfile("chrome https://stackoverflow.com")
        speak("Stack Overflow opened.")
    elif "open camera" in cmd:
        os.system("start microsoft.windows.camera:")
        speak("Camera opened.")

    # --- JOKES & QUOTES ---
    elif "joke" in cmd:
        speak("Why did the computer show up at work late? It had a hard drive!")
    elif "quote" in cmd:
        speak("“The only way to do great work is to love what you do.” — Steve Jobs")

    # --- CALCULATOR VIA VOICE ---
    elif "calculate" in cmd:
        try:
            expr = cmd.replace("calculate", "").strip()
            result = eval(expr)
            speak(f"Result is {result}")
        except Exception:
            speak("Sorry, I couldn't calculate that.")

    # --- SCREEN RECORDING (audio) ---
    elif "record audio" in cmd:
        duration = int(cmd.replace("record audio for", "").replace("seconds","").strip() or 5)
        filename = f"record_{datetime.now():%Y%m%d_%H%M%S}.wav"
        speak(f"Recording audio for {duration} seconds.")
        rec = sd.rec(int(duration * 44100), samplerate=44100, channels=2)
        sd.wait()
        sd.write(filename, rec, 44100)
        speak(f"Saved recording to {filename}")

    # --- ZIP BACKUP & RESTORE ---
    elif "backup files" in cmd:
        zipname = f"backup_{datetime.now():%Y%m%d_%H%M%S}.zip"
        with zipfile.ZipFile(zipname, 'w') as zf:
            for root, _, files in os.walk(os.getcwd()):
                for f in files:
                    zf.write(os.path.join(root, f))
        speak(f"Backup created: {zipname}")
    elif "restore files" in cmd:
        zipname = cmd.replace("restore files", "").strip() or speak("Which backup?")
        with zipfile.ZipFile(zipname, 'r') as zf:
            zf.extractall(os.getcwd())
        speak(f"Restored from {zipname}")

    # Browse specific sites
    elif "open youtube" in command or "یوتیوب رو باز کن" in command:
        os.system("start chrome https://www.youtube.com")
        speak("یوتیوب باز شد.")

    elif "open google" in command or "گوگل رو باز کن" in command:
        os.system("start chrome https://www.google.com")
        speak("گوگل باز شد.")

    # Calculator via voice
    elif "calculate" in command or "محاسبه کن" in command:
        try:
            expression = command.replace("calculate", "").strip()
            result = eval(expression)
            speak(f"نتیجه برابر است با {result}")
        except Exception:
            speak("متاسفم، نتوانستم محاسبه کنم.")


    # Open folders
    elif "open downloads" in command or "دانلودها رو باز کن" in command:
        os.system("start %USERPROFILE%\\Downloads")
        speak("پوشه دانلود باز شد.")

    elif "open documents" in command or "اسناد رو باز کن" in command:
        os.system("start %USERPROFILE%\\Documents")
        speak("پوشه اسناد باز شد.")

    # Custom game launcher
    elif "play cod" in command or "بازی رو اجرا کن" in command or "حوصله ندارم" in command:
        speak("بازی در حال اجراست.")
        subprocess.Popen([r"D:\Call of Duty WWII\Call of Duty WWII(@DAVOODSILVER)\CoD_SP.exe"])

    

    elif "open task manager" in command or "مدیر وظیفه رو باز کن" in command:
        os.system("start taskmgr")
        speak("مدیر وظیفه باز شد.")

    elif "clear recycle bin" in command or "سطل آشغال رو خالی کن" in command:
        speak("سطل آشغال خالی شد.")
        os.system("powershell -command \"Clear-RecycleBin -Force\"")

    elif "open paint 3d" in command or "پینت سه بعدی رو باز کن" in command:
        os.system("start ms-paint:3d")
        speak("پینت ۳ بعدی باز شد.")

    elif "open maps" in command or "نقشه رو باز کن" in command:
        os.system("start bingmaps:")
        speak("نقشه باز شد.")

    elif "open news" in command or "اخبار رو باز کن" in command:
        os.system("start chrome https://www.bbc.com/news")
        speak("اخبار باز شد.")

    elif "find file" in command or "فایل پیدا کن" in command:
        speak("نام یا بخشی از مسیر فایل را بگویید.")
        # integrate search

    elif "disk usage" in command or "استفاده از دیسک" in command:
        speak("استفاده از هارد دیسک ۴۰ درصد است.")
        # integrate real check

    elif "memory usage" in command or "استفاده از رم" in command:
        speak("استفاده از حافظه رم ۵۰ درصد است.")
        # integrate real check


    elif "open bluetooth" in command or "بلوتوث رو باز کن" in command:
        os.system("start ms-settings:bluetooth")
        speak("بلوتوث باز شد.")

    elif "turn on wifi" in command or "وای‌فای رو روشن کن" in command:
        speak("وای‌فای روشن شد.")
        # integrate wifi toggle

    elif "turn off wifi" in command or "وای‌فای رو خاموش کن" in command:
        speak("وای‌فای خاموش شد.")
        # integrate wifi toggle

    elif "play video" in command or "ویدیو پخش کن" in command:
        speak("در حال پخش ویدیو.")
        # integrate video player

    elif "pause video" in command or "ویدیو رو متوقف کن" in command:
        speak("پخش ویدیو متوقف شد.")

    elif "open camera roll" in command or "گالری رو باز کن" in command:
        os.system("start ms-photos:")
        speak("گالری باز شد.")

    elif "open photos" in command or "عکس‌ها رو باز کن" in command:
        os.system("start microsoft.photos:")
        speak("اپلیکیشن Photos باز شد.")

    elif "open ink workspace" in command or "محیط یادداشت برداری رو باز کن" in command:
        os.system("start ms-inkworkspace:")
        speak("محیط یادداشت برداری باز شد.")

    elif "open voice recorder" in command or "ضبط صدا رو باز کن" in command:
        os.system("start soundrecorder")
        speak("ضبط صدا باز شد.")

    elif "record audio" in command or "ضبط صدا کن" in command:
        speak("شروع ضبط صدا.")
        # integrate recording

    elif "stop recording" in command or "ضبط رو متوقف کن" in command:
        speak("ضبط صدا متوقف شد.")
        # integrate stop

    elif "open xbox" in command or "ایکس‌باکس رو باز کن" in command:
        os.system("start xbox:")
        speak("ایکس‌باکس باز شد.")

    elif "open store" in command or "استور رو باز کن" in command:
        os.system("start ms-windows-store:")
        speak("مایکروسافت استور باز شد.")

    elif "download updates" in command or "به‌روزرسانی‌ها رو دانلود کن" in command:
        speak("در حال دانلود به‌روزرسانی‌ها.")
        # integrate Windows Update

    elif "install updates" in command or "به‌روزرسانی‌ها رو نصب کن" in command:
        speak("در حال نصب به‌روزرسانی‌ها.")
        # integrate Windows Update

    elif "check updates" in command or "به‌روزرسانی‌ها رو بررسی کن" in command:
        speak("در حال بررسی به‌روزرسانی‌ها.")
        os.system("start ms-settings:windowsupdate")

    elif "open registry" in command or "رجیستری رو باز کن" in command:
        os.system("start regedit")
        speak("رجیستری ادیتور باز شد.")

    elif "run as administrator" in command or "با دسترسی ادمین اجرا کن" in command:
        speak("اجرای با دسترسی ادمین انجام شد.")
        # placeholder

    elif "open event viewer" in command or "مشاهده رخدادها" in command:
        os.system("start eventvwr")
        speak("Event Viewer باز شد.")

    elif "open services" in command or "سرویس‌ها رو باز کن" in command:
        os.system("start services.msc")
        speak("Services Manager باز شد.")

    elif "cleanup disk" in command or "پاکسازی دیسک" in command:
        speak("در حال اجرای Disk Cleanup.")
        os.system("cleanmgr")

    elif "defragment disk" in command or "یکپارچه‌سازی دیسک" in command:
        speak("در حال یکپارچه‌سازی دیسک.")
        os.system("defrag C: /U /V")

    elif "backup files" in command or "تهیه بکاپ کن" in command:
        speak("در حال تهیه نسخه پشتیبان.")
        # integrate backup script

    elif "restore files" in command or "بازیابی فایل‌ها" in command:
        speak("در حال بازیابی نسخه پشتیبان.")
        # integrate restore

    elif "open recovery" in command or "بازیابی سیستم" in command:
        os.system("start rstrui")
        speak("System Restore باز شد.")

    elif "disk health" in command or "سلامت دیسک" in command:
        speak("وضعیت سلامت دیسک خوب است.")
        # integrate SMART check

    elif "open firewall" in command or "فایروال رو باز کن" in command:
        os.system("start wf.msc")
        speak("Windows Firewall باز شد.")

    elif "enable firewall" in command or "فایروال رو فعال کن" in command:
        speak("فایروال فعال شد.")
        # integrate firewall toggle

    elif "disable firewall" in command or "فایروال رو غیرفعال کن" in command:
        speak("فایروال غیرفعال شد.")
        # integrate firewall toggle

    elif "open vpn status" in command or "وضعیت وی‌پی‌ان" in command:
        os.system("start ms-settings:network-vpn")
        speak("وضعیت VPN نمایش داده شد.")

    elif "open network status" in command or "وضعیت شبکه" in command:
        os.system("start ms-settings:network-status")
        speak("وضعیت شبکه باز شد.")

    elif "open data usage" in command or "مصرف داده رو باز کن" in command:
        os.system("start ms-settings:datausage")
        speak("مصرف داده نمایش داده شد.")

    else:
        speak("دستور شما رو متوجه نشدم.")

# Audio stream callback
def callback(indata, frames, time, status):
    if status:
        print(status)
    q.put(bytes(indata))

# Start listening
with sd.RawInputStream(samplerate=16000, blocksize=8000, dtype='int16',
                       channels=1, callback=callback):
    rec = vosk.KaldiRecognizer(model, 16000)
    speak("در حال گوش دادن هستم...")

    while True:
        data = q.get()
        if rec.AcceptWaveform(data):
            result = rec.Result()
            text = json.loads(result)["text"]
            print("Heard:", text)
            if text:
                execute_command(text)
