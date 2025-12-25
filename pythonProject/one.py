from flask import Flask, render_template, request, jsonify
import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import smtplib

app = Flask(__name__)

# Initialize text-to-speech engine
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

# Define your existing functions here (wishMe, takeCommand, etc.)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/ask', methods=['POST'])
def ask():
    user_input = request.json['query']
    # Process the user input as per your existing logic
    # For example, if 'wikipedia' in user_input, call Wikipedia function
    # Return the response back to the frontend
    return jsonify(response="Your response here")

if __name__ == '__main__':
    app.run(debug=True)
