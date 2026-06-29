#include <Arduino.h>

const uint8_t  LED_PIN     = LED_BUILTIN;
const uint8_t  SENSOR_PIN  = A0;
const uint16_t MAX_SAMPLES = 2000;

uint16_t sensorArray[MAX_SAMPLES];
volatile uint16_t sampleCount = 0;
volatile bool takeSample      = false;
volatile bool recordingDone   = false;

void TCC1_Handler()
{
  TCC1->INTFLAG.bit.MC0 = 1;
  if (sampleCount < MAX_SAMPLES)
    takeSample = true;
  else
    recordingDone = true;
}

void startTimer()
{
  PM->APBCMASK.reg |= PM_APBCMASK_TCC1;
  GCLK->CLKCTRL.reg = GCLK_CLKCTRL_ID(GCM_TCC0_TCC1) |
                      GCLK_CLKCTRL_GEN_GCLK0           |
                      GCLK_CLKCTRL_CLKEN;
  while (GCLK->STATUS.bit.SYNCBUSY);

  TCC1->CTRLA.reg = TCC_CTRLA_SWRST;
  while (TCC1->SYNCBUSY.bit.SWRST);

  TCC1->CTRLA.reg  = TCC_CTRLA_PRESCALER_DIV1024;
  TCC1->WAVE.reg   = TCC_WAVE_WAVEGEN_MFRQ;
  while (TCC1->SYNCBUSY.bit.WAVE);
  TCC1->CC[0].reg  = 46;
  while (TCC1->SYNCBUSY.bit.CC0);

  TCC1->INTENSET.reg = TCC_INTENSET_MC0;
  NVIC_EnableIRQ(TCC1_IRQn);

  TCC1->CTRLA.reg |= TCC_CTRLA_ENABLE;
  while (TCC1->SYNCBUSY.bit.ENABLE);
}

void stopTimer()
{
  TCC1->CTRLA.reg &= ~TCC_CTRLA_ENABLE;
  while (TCC1->SYNCBUSY.bit.ENABLE);
}

void waitForStart()
{
  // Reset sample state
  sampleCount   = 0;
  takeSample    = false;
  recordingDone = false;

  digitalWrite(LED_PIN, LOW);
  Serial.flush();

  // Wait for START command
  while (true) {
    if (Serial.available()) {
      String cmd = Serial.readStringUntil('\n');
      cmd.trim();
      if (cmd == "START") {
        Serial.println("READY");
        break;
      }
    }
  }

  // Begin recording
  digitalWrite(LED_PIN, HIGH);
  startTimer();
}

void setup()
{
  Serial.begin(115200);
  pinMode(LED_PIN, OUTPUT);
  analogRead(SENSOR_PIN); // warm up ADC
  waitForStart();
}

void loop()
{
  if (takeSample)
  {
    takeSample = false;
    sensorArray[sampleCount] = analogRead(SENSOR_PIN);
    sampleCount++;
  }

  if (recordingDone)
  {
    stopTimer();
    recordingDone = false;
    digitalWrite(LED_PIN, LOW);

    // Send all samples
    for (uint16_t i = 0; i < MAX_SAMPLES; i++)
      Serial.println(sensorArray[i]);
    Serial.println("--- END OF DATA ---");

    // Wait for next START command automatically
    waitForStart();
  }
}