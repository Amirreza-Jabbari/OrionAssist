## OrionAssist

OrionAssist is a powerful, extensible voice-activated assistant for Windows, built in Python. It leverages VOSK for offline speech recognition, pyttsx3 for text-to-speech, and a range of system-level libraries (psutil, PIL, pycaw, schedule, etc.) to perform real-world tasks via natural voice commands.

---

### 🚀 Features

* **System Controls**

  * Open browser, lock PC, log off, shutdown, restart, hibernate
* **Resource Monitoring**

  * Battery level, CPU/RAM usage, disk usage
* **Time & Date**

  * Speak current time & date
* **Screenshots**

  * Capture and save desktop screenshots
* **Volume Management**

  * Increase, decrease, mute, and unmute system volume
* **Network Controls**

  * Enable/disable Wi-Fi and connect/disconnect VPN
* **Weather Lookup**

  * Fetch live weather using OpenWeatherMap API
* **Alarms & Reminders**

  * Schedule daily alarms and one-off reminders
* **File Operations**

  * Create, find, and delete files with glob patterns
* **Media Playback**

  * Play, pause, and skip tracks via VLC & NirCmd
* **Application Launchers**

  * Open Notepad, Calculator, Paint, Office suite, browsers, and more
* **Jokes & Quotes**

  * Light-hearted responses for “joke” and “quote” commands
* **Calculations**

  * Perform basic arithmetic via voice
* **Audio Recording**

  * Record microphone input for a specified duration
* **ZIP Backup & Restore**

  * Archive current directory; restore from existing ZIP

---

### 📋 Requirements

* **Python 3.8+**
* **Windows 10/11**
* **Third-party packages**

  ```bash
  pip install vosk sounddevice pyttsx3 psutil Pillow requests schedule pycaw comtypes
  ```
* **Optional utilities**

  * **NirCmd** (for media key handling)
  * **VLC** (for media playback)
* **OpenWeatherMap API key**
  Sign up at [https://openweathermap.org](https://openweathermap.org) and set `OWM_API_KEY` in the script.

---

### ⚙️ Installation

1. **Clone the repo**

   ```bash
   git clone https://github.com/Amirreza-Jabbari/OrionAssist.git
   cd OrionAssist
   ```

2. **Install dependencies**

   ```bash
   pip install -r requirements.txt
   ```

3. **Download VOSK model**

   * Place the unpacked model folder at `vosk-model-small-fa-0.42` (or update `MODEL_PATH`).
   * Russian, English, or other language models work too—just adjust path.

4. **Configure the script**

   * Open `orionassist.py`.
   * Set `MODEL_PATH`, `OWM_API_KEY`, `MUSIC_PATH`, and your VPN credentials.

---

### ▶️ Usage

Run the assistant script:

```bash
python orionassist.py
```

Speak one of the supported commands clearly. For example:

* “Hey Orion, open browser”
* “What is the CPU usage?”
* “Set alarm at 07:30”
* “What’s the weather in London?”
* “Create file notes.txt”
* “Play song”
* “Record audio for 10 seconds”

The assistant will confirm via voice and execute the task.

---

### 🛠️ Architecture

* **Speech recognition**: VOSK offline model
* **TTS**: pyttsx3
* **Audio I/O**: sounddevice
* **System stats**: psutil
* **Screenshots**: PIL.ImageGrab
* **Volume control**: Windows Core Audio API (pycaw)
* **Scheduling**: schedule (background thread)
* **Network**: `netsh` & `rasdial`
* **Weather**: OpenWeatherMap REST API
* **File operations**: Python’s `glob` & `zipfile`

---

### 🤝 Contributing

1. Fork the repo
2. Create a feature branch (`git checkout -b feature/my-feature`)
3. Commit your changes (`git commit -m "Add awesome feature"`)
4. Push to your branch (`git push origin feature/my-feature`)
5. Open a Pull Request

Please adhere to PEP8, include docstrings, and write tests for new functionality.

---

### 📄 License

This project is released under the **MIT License**. See [LICENSE](./LICENSE) for details.
