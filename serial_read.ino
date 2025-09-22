
String lastInput = ""; // Declare a string to store the last user input

void setup() {
  Serial.begin(9600); // Initialize serial communication at 9600 baud rate
  pinMode(13, OUTPUT); // Set pin 13 as an output
  pinMode(12,OUTPUT);
  pinMode(11,OUTPUT);
  pinMode(10,OUTPUT);
  digitalWrite(13, LOW); // Initialize pin 13 to LOW
  digitalWrite(12, LOW); // Initialize pin 12 to LOW
  digitalWrite(11, LOW);
  digitalWrite(10, LOW);
}

void loop() {
  ui_input();
  Serial.println(lastInput);
}
void ledOn(int ledPin){//open a certain led only
  killAll();
  digitalWrite(ledPin, HIGH);
}
void killAll(){//kill all the leds
  digitalWrite(13, LOW);
  digitalWrite(12, LOW);
  digitalWrite(11, LOW);
  digitalWrite(10, LOW);
}

void ui_input(){
  if (Serial.available() > 0) { // Check if data is available to read
    String input = Serial.readString(); // Read the incoming string
    input.trim(); // Remove any trailing whitespace or newline
    if (input!= ""){
      lastInput = input; // Update the lastInput with new input
      ledOn(11);//RED will show you there is a input change
    }
  } else if (lastInput != "") { // If no new input and lastInput is not empty
    //do nothing
  }
  delay(200); // Small delay to avoid flooding the serial output
  if (lastInput.indexOf("start") != -1) {//checking if the input is start
      ledOn(13);//BLUE
      delay(2000);
    } else if (lastInput.indexOf("stop") != -1) {//checking if the input is stop
      ledOn(12);//GREEN
      delay(2000);
    }
}