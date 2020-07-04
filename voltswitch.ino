/* Basic Raw HID Example
   Teensy can send/receive 64 byte packets with a
   dedicated program running on a PC or Mac.

   You must select Raw HID from the "Tools > USB Type" menu

   Optional: LEDs should be connected to pins 0-7,
   and analog signals to the analog inputs.

   This example code is in the public domain.
*/

// Voltage switch states.
// The order is such that 5V is default even without VCC.
// Technically 7V is possible when 
const char fiveV = 0;
const char twelveV = 1;
const char zeroV = 2;

// Pins to be used
const int ledPin = 13;
const int v12Pin = 23;
const int v5Pin = 22;
const int v0Pin = 21;
const int relay1 = 16;
const int relay2 = 17;

void setup() {
  Serial.begin(9600);
  Serial.println(F("RawHID Example"));
  
  // Set the relay controller pin modes
  pinMode(relay1, OUTPUT);
  pinMode(relay2, OUTPUT);
  digitalWrite(relay1, LOW);
  digitalWrite(relay2, LOW);

  // Set the override pin modes
  pinMode(v12Pin, INPUT_PULLUP);
  pinMode(v5Pin, INPUT_PULLUP);
  pinMode(v0Pin, INPUT_PULLUP);

  // Pin 13 is the onboard LED just to signal power on
  pinMode(ledPin, OUTPUT);
  digitalWrite(ledPin, HIGH);
}

// RawHID packets are always 64 bytes
byte buffer[64];
elapsedMillis msUntilNextSend;
unsigned int packetCount = 0;

// State stores the current relay state
char state = 0;

void loop() {
  int n;
  char newState = 0;

  // If v12Pin or v0Pin are jumpered, they override HID
  if (!digitalRead(v12Pin) && (state != twelveV)) {
    Serial.println("Doing 12V");
    digitalWrite(relay1, HIGH);
    digitalWrite(relay2, LOW);
    state = twelveV;
  } else if (!digitalRead(v0Pin) && (state != zeroV)) {
    Serial.println("Doing 0V");
    digitalWrite(relay1, LOW);
    digitalWrite(relay2, HIGH);
    state = zeroV;
  } else if (!digitalRead(v5Pin) && (state != fiveV)) {
    Serial.println("Doing 5V");
    digitalWrite(relay1, LOW);
    digitalWrite(relay2, LOW);
    state = fiveV;
  } else {
    // Timeout in millis: hardware/teensy/avr/cores/teensy3/usb_rawhid.c
    n = RawHID.recv(buffer, 0); // 0 timeout = do not wait
    if (n > 0) {
      // We only care about 2 bits of the 64-byte payload
      Serial.print(F("Received packet, first byte: "));
      Serial.println((int)buffer[0]);
  
      newState = buffer[0] & 0x3;
      if (newState != state) {
        switch(newState) {
          case zeroV:
            Serial.println("Doing HID 0V");
            digitalWrite(relay1, LOW);
            digitalWrite(relay2, HIGH);
            state = newState;
            break;
          case fiveV:
            Serial.println("Doing HID 5V");
            digitalWrite(relay1, LOW);
            digitalWrite(relay2, LOW);
            state = newState;
            break;
          case twelveV:
            Serial.println("Doing HID 12V");
            digitalWrite(relay1, HIGH);
            digitalWrite(relay2, LOW);
            state = newState;
            break;
        }
      }
    }
    
    // Every 2 seconds, send the current state back to computer
    if (msUntilNextSend > 2000) {
      msUntilNextSend = msUntilNextSend - 2000;
      // Set the buffer to just our current state
      buffer[0] = state;
      
      // fill the rest with zeros
      for (int i=1; i<64; i++) {
        buffer[i] = 0;
      }
      
      // actually send the packet
      n = RawHID.send(buffer, 100);
      if (n > 0) {
        Serial.print(F("Transmit packet "));
        Serial.println(packetCount);
        packetCount = packetCount + 1;
      } else {
        Serial.println(F("Unable to transmit packet"));
      }
    }
  }
}
