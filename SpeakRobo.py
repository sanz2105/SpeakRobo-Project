import win32com.client as wincom
import tkinter as tk
from tkinter import ttk

def speak_text(text, voice):
    speak = wincom.Dispatch("SAPI.SpVoice")
    speak.Voice = speak.GetVoices().Item(voice)
    speak.Speak(text)

def on_submit():
    user_input = entry.get()
    selected_voice = voice_combobox.current()

    if user_input.upper() == "BYE":
        print("Goodbye!")
        root.destroy()
    else:
        speak_text(user_input, selected_voice)

# Initialize the tkinter window
root = tk.Tk()
root.title("RoboSpeaker - Created by Sanz")

# Create and place widgets
label = tk.Label(root, text="Enter what you want me to say:")
label.pack(pady=10)

entry = tk.Entry(root, width=50)
entry.pack(pady=10)

voice_label = tk.Label(root, text="Select voice:")
voice_label.pack(pady=5)

voices = [voice.GetDescription() for voice in wincom.Dispatch("SAPI.SpVoice").GetVoices()]
voice_combobox = ttk.Combobox(root, values=voices, state="readonly")
voice_combobox.current(0)  # Default to the first voice
voice_combobox.pack(pady=10)

submit_button = tk.Button(root, text="Submit", command=on_submit)
submit_button.pack(pady=10)

# Start the tkinter event loop
root.mainloop()
