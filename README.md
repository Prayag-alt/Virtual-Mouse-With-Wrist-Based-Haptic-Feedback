# Gesture-Controlled Virtual Mouse with Wrist Haptic Feedback

> Touchless computer control using hand gestures via webcam — with physical 
> haptic confirmation on your wrist. Built with MediaPipe, OpenCV, Win32 API, 
> and ESP32.

---

## Demo

| Default Mode | PowerPoint Mode |
|---|---|
| Right hand moves cursor | Cursor frozen, gestures control slides |
| Left hand clicks & scrolls | Index = next, Peace = prev, Palm = zoom |
| Fist = freeze cursor | Three fingers = start slideshow |

---

## Features

### Default Mode
- Real-time cursor control via right hand index finger
- Left click — thumb + index pinch
- Right click — thumb + middle pinch  
- Double click — all fingers pinch
- Scroll — index + middle extended, move up/down
- Cursor freeze — make a fist

### PowerPoint Mode
- All default gestures completely disabled — zero bleed
- Cursor frozen at last position
- Next slide — index finger only
- Previous slide — peace sign
- Start slideshow — three fingers held 1s
- Stop slideshow — fist held 1.5s
- Zoom in/out — open palm + move hand up/down
- Freeze/unfreeze cursor — fist held 0.8s
- Floating toolbar — both index fingers up 0.3s
- Enter/exit PPT mode — both fists held 2s

### Hardware Haptic Feedback
- ESP32 wrist device with coin vibration motor
- Distinct buzz patterns per gesture action
- WiFi UDP communication from PC to wrist
- Powered by 3.7V LiPo with TP4056 safe charging

---

## Architecture
```
Camera Thread (60 FPS)
      ↓
MediaPipe HandLandmarker (every 2nd frame)
      ↓
Gesture Engine (Kalman filter + state machines)
      ↓
Mode Router ──→ Default Mode (cursor + clicks)
            └──→ PPT Mode (frozen cursor + gestures)
                      ↓
              Win32 ctypes output (~0.1ms latency)
                      ↓
              ESP32 UDP haptic feedback
```

---

## Performance

| Metric | Value |
|---|---|
| Total pipeline latency | < 100ms |
| Mouse event latency | ~0.1ms (Win32 ctypes) |
| Gesture processing rate | 30 FPS effective |
| Static gesture accuracy | 99%+ |
| CPU saving (frame skip) | ~50% |

---

## Tech Stack

| Component | Technology |
|---|---|
| Language | Python 3.x |
| Hand tracking | MediaPipe HandLandmarker |
| Camera pipeline | OpenCV (threaded, CAP_DSHOW) |
| Cursor smoothing | 2D Kalman filter (NumPy) |
| Mouse/keyboard output | ctypes Win32 user32 |
| Toolbar overlay | Tkinter (topmost transparent window) |
| Haptic firmware | ESP32 + Arduino C++ |
| Power system | TP4056 + LiPo 3.7V 400mAh |

---

## Hardware Components

| Component | Specification |
|---|---|
| Microcontroller | ESP32 DevKit v1 |
| Vibration motor | Coin motor 3V (1020/1027) |
| Transistor | BC547 NPN |
| Protection diode | 1N4007 |
| Base resistor | 100Ω |
| Battery | LiPo 3.7V 400mAh |
| Charger | TP4056 with protection (USB-C) |

---

## Circuit — Wrist Haptic Device
```
ESP32 GPIO15 ──→ 100Ω ──→ BC547 Base
ESP32 3.3V   ──→ Motor (+) red wire
Motor (-)  blue wire ──→ BC547 Collector
BC547 Emitter ──→ ESP32 GND
1N4007 diode across motor (cathode to + side)
LiPo (+) ──→ TP4056 BAT+
LiPo (−) ──→ TP4056 BAT−
TP4056 OUT+ ──→ ESP32 VIN
TP4056 OUT− ──→ ESP32 GND
```

---

## Installation
```bash
# Clone the repo
git clone https://github.com/yourusername/gesture-virtual-mouse.git
cd gesture-virtual-mouse

# Create virtual environment
python -m venv venv
venv\Scripts\activate  # Windows

# Install dependencies
pip install -r requirements.txt

# Download MediaPipe hand landmark model
# Place hand_landmarker.task in project root
# Download from: https://ai.google.dev/edge/mediapipe/solutions/vision/hand_landmarker

# Run
python main.py
```

---

## Requirements
```
mediapipe
opencv-python
numpy
```

> No pyautogui. All mouse and keyboard control is handled directly 
> via Win32 ctypes for ~0.1ms latency.

---

## Project Structure
```
gesture-virtual-mouse/
├── main.py                  # Main pipeline
├── hand_landmarker.task     # MediaPipe model (download separately)
├── requirements.txt
├── esp32/
│   └── haptic_firmware.ino  # ESP32 Arduino firmware
└── README.md
```

---

## Key Engineering Decisions

**Why Win32 ctypes instead of pyautogui?**
pyautogui averages 15–20ms per mouse event. Win32 ctypes talks directly to the OS — 0.1ms. 150× faster.

**Why Kalman filter instead of EMA?**
EMA just averages recent positions — it doesn't understand velocity. The Kalman filter uses a constant-velocity motion model so it predicts where your hand is going between frames, eliminating jitter spikes from micro-tremors without adding lag.

**Why frame-skip MediaPipe?**
Running MediaPipe every frame at 60 FPS saturates the CPU. Running it every 2nd frame halves the inference load while the cursor still updates every frame using the cached result — no perceived difference in responsiveness.

**Why separate mode branches?**
In PPT mode, zero default gesture code runs. The two modes are completely separate if/else branches. This makes it physically impossible for cursor movement or click events to bleed into a presentation accidentally.

---

**Guide:** Dr. Priyang Bhatt
**Institution:** G H Patel College of Engineering & Technology, CVM University
**Department:** Computer Engineering (CSD)
**Course:** Mini Project — 202040601

---

## License

MIT License — free to use, modify, and distribute with attribution.
