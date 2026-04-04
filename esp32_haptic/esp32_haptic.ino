/*
 * ═══════════════════════════════════════════════════════════════
 *  ESP32  Haptic Feedback Controller  —  Virtual Mouse / PPT
 * ═══════════════════════════════════════════════════════════════
 *
 *  Hardware
 *  --------
 *    ESP32 DevKit v1 (or any ESP32 board)
 *    Coin vibration motor   → GPIO 13   (via NPN transistor / MOSFET)
 *    TP4056 charging module → LiPo battery → 3.3 V regulator (or
 *                             direct on VBAT if your board supports it)
 *    Optional: LED on GPIO 2 (on-board) for status
 *
 *  Wiring (motor)
 *  ──────────────
 *    ESP32  GPIO 13  ──▶  Base of NPN (2N2222 / BC547) via 1 kΩ
 *    Collector  ──▶  Motor (–)
 *    Motor (+)  ──▶  3.3 V / VBAT
 *    Emitter    ──▶  GND
 *    Flyback diode (1N4148) across motor terminals (cathode to +)
 *
 *  Communication
 *  ─────────────
 *    USB Serial  @ 115200 baud   (or Bluetooth Serial — see below)
 *    Single-byte command protocol:
 *
 *    ┌──────────┬───────────────────────────────────────────┐
 *    │  Byte    │  Event                                    │
 *    ├──────────┼───────────────────────────────────────────┤
 *    │  'L'     │  Left click                               │
 *    │  'R'     │  Right click                              │
 *    │  'D'     │  Double click                             │
 *    │  'S'     │  Scroll tick                              │
 *    │  'P'     │  PPT mode ENTER                           │
 *    │  'p'     │  PPT mode EXIT                            │
 *    │  'N'     │  Next slide                               │
 *    │  'B'     │  Previous (Back) slide                    │
 *    │  'Z'     │  Zoom in / out tick                       │
 *    │  'F'     │  Slideshow started (F5)                   │
 *    │  'E'     │  Slideshow stopped (Esc)                  │
 *    │  'T'     │  Toolbar button selected                  │
 *    │  'C'     │  Cursor frozen                            │
 *    │  'c'     │  Cursor unfrozen                          │
 *    │  'H'     │  Heartbeat (keep-alive, gentle nudge)     │
 *    └──────────┴───────────────────────────────────────────┘
 */

// ── Uncomment ONE of these to choose your serial transport ──
#define USE_USB_SERIAL        // ← default: communicate over USB
// #define USE_BT_SERIAL      // ← uncomment for Classic Bluetooth SPP

#ifdef USE_BT_SERIAL
  #include "BluetoothSerial.h"
  BluetoothSerial BTSerial;
  #define COM  BTSerial
#else
  #define COM  Serial
#endif

// ─────────────────────────────────────────────────────────
//  Pin definitions
// ─────────────────────────────────────────────────────────
#define MOTOR_PIN     13      // PWM-capable GPIO for vibration motor
#define LED_PIN        2      // on-board LED (status indicator)
#define PWM_CHANNEL    0
#define PWM_FREQ    5000      // 5 kHz — inaudible for motor drive
#define PWM_RES        8      // 8-bit resolution: 0-255

// ─────────────────────────────────────────────────────────
//  Vibration intensity presets (0-255)
// ─────────────────────────────────────────────────────────
#define INTENSITY_SOFT      120
#define INTENSITY_MEDIUM    180
#define INTENSITY_STRONG    240
#define INTENSITY_MAX       255

// ─────────────────────────────────────────────────────────
//  Forward declarations
// ─────────────────────────────────────────────────────────
void vibrateMotor(uint8_t intensity, uint16_t durationMs);
void vibratePattern(uint8_t intensity, uint16_t onMs, uint16_t offMs, uint8_t repeats);
void vibrateRamp(uint8_t startIntensity, uint8_t endIntensity, uint16_t durationMs);

// ─────────────────────────────────────────────────────────
//  Haptic patterns — each event gets a unique feel
// ─────────────────────────────────────────────────────────

// Left click — single short, crisp tap
void haptic_left_click() {
  vibrateMotor(INTENSITY_STRONG, 50);
}

// Right click — two quick soft pulses
void haptic_right_click() {
  vibratePattern(INTENSITY_MEDIUM, 40, 50, 2);
}

// Double click — three rapid taps
void haptic_double_click() {
  vibratePattern(INTENSITY_STRONG, 35, 40, 3);
}

// Scroll tick — very gentle short buzz
void haptic_scroll() {
  vibrateMotor(INTENSITY_SOFT, 25);
}

// PPT mode ENTER — ascending ramp  (feels like "powering up")
void haptic_ppt_enter() {
  vibrateRamp(60, INTENSITY_MAX, 400);
  delay(80);
  vibrateMotor(INTENSITY_MAX, 150);
}

// PPT mode EXIT — descending ramp  (feels like "powering down")
void haptic_ppt_exit() {
  vibrateRamp(INTENSITY_MAX, 40, 400);
}

// Next slide — swift forward nudge
void haptic_next_slide() {
  vibrateMotor(INTENSITY_MEDIUM, 60);
  delay(40);
  vibrateMotor(INTENSITY_STRONG, 40);
}

// Previous slide — swift backward nudge (reverse of next)
void haptic_prev_slide() {
  vibrateMotor(INTENSITY_STRONG, 40);
  delay(40);
  vibrateMotor(INTENSITY_MEDIUM, 60);
}

// Zoom tick — medium pulse
void haptic_zoom() {
  vibrateMotor(INTENSITY_MEDIUM, 45);
}

// Slideshow started — triumphant double buzz
void haptic_slideshow_start() {
  vibrateMotor(INTENSITY_STRONG, 100);
  delay(80);
  vibrateMotor(INTENSITY_MAX, 150);
}

// Slideshow stopped — single firm stop
void haptic_slideshow_stop() {
  vibrateMotor(INTENSITY_MAX, 200);
}

// Toolbar button selected — click-like confirmation
void haptic_toolbar_select() {
  vibrateMotor(INTENSITY_MEDIUM, 55);
}

// Cursor frozen — "lock" pattern  (two evenly spaced taps)
void haptic_cursor_frozen() {
  vibratePattern(INTENSITY_STRONG, 60, 60, 2);
}

// Cursor unfrozen — single releasing pulse
void haptic_cursor_unfrozen() {
  vibrateMotor(INTENSITY_SOFT, 80);
}

// Heartbeat — very subtle alive-check (default mode idle)
void haptic_heartbeat() {
  vibrateMotor(INTENSITY_SOFT, 20);
  delay(120);
  vibrateMotor(INTENSITY_SOFT, 20);
}


// ═════════════════════════════════════════════════════════
//  setup()
// ═════════════════════════════════════════════════════════
void setup() {
  // Configure motor PWM
  ledcSetup(PWM_CHANNEL, PWM_FREQ, PWM_RES);
  ledcAttachPin(MOTOR_PIN, PWM_CHANNEL);
  ledcWrite(PWM_CHANNEL, 0);

  // Status LED
  pinMode(LED_PIN, OUTPUT);
  digitalWrite(LED_PIN, LOW);

  // Serial
  #ifdef USE_BT_SERIAL
    BTSerial.begin("VirtualMouse_Haptic");   // Bluetooth device name
    Serial.begin(115200);
    Serial.println("[Haptic] Bluetooth SPP started — device: VirtualMouse_Haptic");
  #else
    Serial.begin(115200);
    Serial.println("[Haptic] USB Serial ready @ 115200");
  #endif

  // Boot-up feedback — brief confirm vibration
  vibrateMotor(INTENSITY_MEDIUM, 100);
  delay(100);
  vibrateMotor(INTENSITY_MEDIUM, 100);
  digitalWrite(LED_PIN, HIGH);

  Serial.println("[Haptic] System initialised. Waiting for commands...");
}


// ═════════════════════════════════════════════════════════
//  loop()
// ═════════════════════════════════════════════════════════
void loop() {
  if (COM.available()) {
    char cmd = (char)COM.read();

    // Blink LED on activity
    digitalWrite(LED_PIN, LOW);

    switch (cmd) {
      case 'L':  haptic_left_click();       break;
      case 'R':  haptic_right_click();       break;
      case 'D':  haptic_double_click();      break;
      case 'S':  haptic_scroll();            break;
      case 'P':  haptic_ppt_enter();         break;
      case 'p':  haptic_ppt_exit();          break;
      case 'N':  haptic_next_slide();        break;
      case 'B':  haptic_prev_slide();        break;
      case 'Z':  haptic_zoom();              break;
      case 'F':  haptic_slideshow_start();   break;
      case 'E':  haptic_slideshow_stop();    break;
      case 'T':  haptic_toolbar_select();    break;
      case 'C':  haptic_cursor_frozen();     break;
      case 'c':  haptic_cursor_unfrozen();   break;
      case 'H':  haptic_heartbeat();         break;
      default:
        // Unknown command — ignore
        break;
    }

    digitalWrite(LED_PIN, HIGH);
  }
}


// ═════════════════════════════════════════════════════════
//  Low-level vibration helpers
// ═════════════════════════════════════════════════════════

/**
 * Vibrate at a fixed intensity for a given duration, then stop.
 */
void vibrateMotor(uint8_t intensity, uint16_t durationMs) {
  ledcWrite(PWM_CHANNEL, intensity);
  delay(durationMs);
  ledcWrite(PWM_CHANNEL, 0);
}

/**
 * Repeat an on/off pattern.
 *   intensity — PWM duty (0-255)
 *   onMs      — vibration ON time per pulse
 *   offMs     — gap between pulses
 *   repeats   — number of pulses
 */
void vibratePattern(uint8_t intensity, uint16_t onMs, uint16_t offMs, uint8_t repeats) {
  for (uint8_t i = 0; i < repeats; i++) {
    ledcWrite(PWM_CHANNEL, intensity);
    delay(onMs);
    ledcWrite(PWM_CHANNEL, 0);
    if (i < repeats - 1) delay(offMs);
  }
}

/**
 * Smooth ramp from startIntensity → endIntensity over durationMs.
 * Creates a "power up" or "power down" feeling.
 */
void vibrateRamp(uint8_t startIntensity, uint8_t endIntensity, uint16_t durationMs) {
  const uint16_t steps = 50;
  uint16_t stepDelay = durationMs / steps;
  float stepSize = (float)(endIntensity - startIntensity) / (float)steps;

  for (uint16_t i = 0; i <= steps; i++) {
    uint8_t val = (uint8_t)(startIntensity + stepSize * i);
    ledcWrite(PWM_CHANNEL, val);
    delay(stepDelay);
  }
  ledcWrite(PWM_CHANNEL, 0);
}
