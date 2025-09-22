// Define the pin
const int pin2 = 2;

// Global variable to track pin 2 state
bool pin2State = LOW;

void setup() {
  // Initialize serial communication at 57600 baud (match with Windows app)
  Serial.begin(57600);
  
  // Set pin 2 as output
  pinMode(pin2, OUTPUT);
  
  // Initialize pin 2 to LOW
  digitalWrite(pin2, pin2State);
  
  // Initial serial output prompt
  Serial.println("Serial communication started. Waiting for commands...");
}

void loop() {
  // Check if data is available to read
  //Serial.print("Hello");
  while (Serial.available() > 0) {
    String input = Serial.readString(); // Read all available data
    input.trim(); // Remove any whitespace or newline

    // Echo all input back to serial
    Serial.print("Received: ");
    Serial.println(input);

    // Check for WASD commands
    if (input == ".KEY.W;" || input == ".KEY.A;" || input == ".KEY.S;" || input == ".KEY.D;") {
      // Toggle pin 2 state
      pin2State = !pin2State; // Reverse the state
      digitalWrite(pin2, pin2State); // Update the pin
      Serial.print("Pin 2 toggled to ");
      Serial.print(pin2State ? "HIGH" : "LOW");
      Serial.print(" by: ");
      Serial.println(input);
    }
  }

  // Optional: Add a delay to prevent overwhelming the loop
  delay(10);
}