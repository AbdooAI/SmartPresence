#include <SPI.h>
#include <MFRC522.h>

#define SS_PIN 10
#define RST_PIN 9

MFRC522 mfrc522(SS_PIN, RST_PIN);  

void setup() {
  pinMode(LED_BUILTIN, OUTPUT);
  Serial.begin(9600);  
  SPI.begin();         
  mfrc522.PCD_Init();  
}



void loop() {
  
  if (isCardAvailable()) {
    
    
    for (byte i = 0; i < mfrc522.uid.size; i++) {
      Serial.print(mfrc522.uid.uidByte[i] < 0x10 ? " 0" : " ");
      Serial.print(mfrc522.uid.uidByte[i], HEX);
    }
    Serial.println();

    
    mfrc522.PICC_HaltA();

  
    mfrc522.PCD_StopCrypto1();
  }
  delay(100);
   digitalWrite(LED_BUILTIN, HIGH);
  delay(100);
}

bool isCardAvailable() {
  if (!mfrc522.PICC_IsNewCardPresent()) {
    return false;
  }
  if (!mfrc522.PICC_ReadCardSerial()) {
    return false;
  }
  return true;
}