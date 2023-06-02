# smart-aura-sync
Python voice assistant that controls the RGB colors of ASUS motherboard compatible with Aura Sync system

The "Smart Aura Sync" program is a voice assistant developed in Python that allows you to control the RGB colors of ASUS motherboard compatible with the Aura Sync system. Using the pyttsx3 library for speech synthesis and the speech_recognition library for speech recognition, the assistant is able to receive voice commands to change the colors of the motherboard's RGB lighting.

The voice assistant recognizes color-related commands such as "red", "green", "blue", "yellow", "purple", "cyan", "orange", "pink" and "white". Upon receiving a voice command, the program uses the ASUS Aura SDK to apply the chosen colors to the compatible motherboard lighting devices.

In addition, the program has a shutdown feature, which can be activated through the "computer terminate" command. When exiting, the voice assistant uses the pyttsx3 library to play the message "Closing Smart Aura Sync. Goodbye!" before finishing the run.
