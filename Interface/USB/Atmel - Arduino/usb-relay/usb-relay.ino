#define RELAY1_PIN 23  // Pin za prvi relej
#define RELAY2_PIN 22  // Pin za drugi relej

void setup() {
  // Pokrećemo serijsku komunikaciju
  Serial.begin(115200);
  
  // Inicijalizacija pinova za releje
  pinMode(RELAY1_PIN, OUTPUT);
  pinMode(RELAY2_PIN, OUTPUT);
  
  // Inicijalizacija releja (inicijalno isključeni)
  digitalWrite(RELAY1_PIN, LOW);
  digitalWrite(RELAY2_PIN, LOW);
  
  Serial.println("Arduino ready. Awaiting commands...");
}

void loop() {
  // Ako su podaci dostupni na serijskoj vezi
  if (Serial.available()) {
    String command = Serial.readStringUntil('\n');  // Čita komandu do novog reda
    command.trim();  // Uklanja eventualne praznine sa početka i kraja komande
    
    // Proveravamo komande
    if (command == "HELLO") {
      Serial.println("OK");  // Odgovara sa OK kao potvrda da je Arduino povezan
    }
    else if (command == "R11") {
      digitalWrite(RELAY1_PIN, HIGH);  // Upali prvi relej
      Serial.println("Relay 1 ON");
    }
    else if (command == "R10") {
      digitalWrite(RELAY1_PIN, LOW);   // Ugasi prvi relej
      Serial.println("Relay 1 OFF");
    }
    else if (command == "R21") {
      digitalWrite(RELAY2_PIN, HIGH);  // Upali drugi relej
      Serial.println("Relay 2 ON");
    }
    else if (command == "R20") {
      digitalWrite(RELAY2_PIN, LOW);   // Ugasi drugi relej
      Serial.println("Relay 2 OFF");
    }
    else {
      Serial.println("Invalid command");
    }
  }
}
