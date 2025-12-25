import os
import subprocess
from flask import Flask, render_template, request, jsonify
import pyttsx3
import speech_recognition as sr
from datetime import datetime
import wikipedia
import webbrowser
import threading
from queue import Queue
from googleapiclient.discovery import build
import urllib.parse
import win32com.client
from collections import defaultdict
import re
import cv2
import platform
from textblob import TextBlob  # Import TextBlob for sentiment analysis
import pyautogui
import time
import pyscreeze
import threading

reminders = []
app = Flask(__name__)

# Initialize text-to-speech engine
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

# Create a queue for speech requests
speech_queue = Queue()
is_speaking = False
# Store user chat history and behavior analysis
user_history = []
user_behavior = defaultdict(int)
user_timestamps = defaultdict(list)
keyword_frequency = defaultdict(int)

def speech_worker():
    global is_speaking
    while True:
        text = speech_queue.get()
        if text is None:
            break
        is_speaking = True
        engine.say(text)
        engine.runAndWait()
        is_speaking = False
        speech_queue.task_done()

# Start the speech worker thread
threading.Thread(target=speech_worker, daemon=True).start()

def speak(audio):
    """Function to convert text to speech."""
    speech_queue.put(audio)

def stop_speaking():
    global is_speaking
    if is_speaking:
        engine.stop()
        is_speaking = False
        return "Stopped speaking, master."
    else:
        return "I'm not speaking right now, master."

def take_command():
    """Takes voice input and converts it to text."""
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 2  # Allow more time for speaking
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")
    except Exception:
        print("Could not understand master, please say again...")
        return "None"
    return query.lower()

def browse_topic(topic):
    """Open browser with Google search results for the given topic."""
    encoded_query = urllib.parse.quote_plus(topic)
    search_url = f"https://www.google.com/search?q={encoded_query}"
    webbrowser.open(search_url)
    return f"Opening browser with results for: {topic}"

def is_chrome_running():
    """Check if Google Chrome is running."""
    try:
        output = subprocess.check_output("tasklist", shell=True)
        return "chrome.exe" in output.decode()
    except Exception as e:
        print(f"Error checking Chrome status: {e}")
        return False

def google_search(query):
    """Performs Google Search using Custom Search API."""
    api_key = 'AIzaSyAo8QiTOZAgQnCa9eaFh9E2bZVsfnj8UQA'  # Replace with your API Key
    cse_id = 'a203374f897c743d4'

    service = build("customsearch", "v1", developerKey=api_key)

    try:
        res = service.cse().list(q=query, cx=cse_id).execute()
        if 'items' in res:
            results = [(item['title'], item['link'], item.get('snippet', 'No description available master.')) for item in res['items'][:2]]

            result_message = "Select a site from the following results master:\n" + "\n".join(
                f"{index + 1}. {title}\n   {link}\n   {snippet}\n" for index, (title, link, snippet) in enumerate(results)
            )
            print(result_message)
            return result_message  # Return the result message to be used in the response
    except Exception as e:
        error_msg = f"Search Error: {str(e)}"
        print(error_msg)
        speak("Sorry, I encountered a problem performing that search. Please try again master.")
        return "Search operation failed"

def open_application(app_path):
    """Open an application using the provided path."""
    try:
        # Check if the application exists at the provided path
        if os.path.exists(app_path):
            subprocess.Popen([app_path], shell=True)
            return "Opened application"
        else:
            return f"Application not found at: {app_path}"
    except Exception as e:
        return f"Error opening application: {str(e)}"

def type_into_word(save_path="C:/VoiceDocs/"):
    """Opens Word, types voice input, and saves automatically"""
    try:
        # Initialize components
        recognizer = sr.Recognizer()
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Add()

        # Get voice input
        with sr.Microphone() as source:
            speak("Please speak the text you want to save")
            audio = recognizer.listen(source, timeout=8)

        # Convert speech to text
        text = f"Voice Document\n{datetime.now().strftime('%Y-%m-%d %H:%M')}\n\n{recognizer.recognize_google(audio)}"

        # Insert text and save
        doc.Content.Text = text
        filename = f"VoiceDoc_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        full_path = os.path.join(save_path, filename)

        # Create directory if needed
        os.makedirs(save_path, exist_ok=True)

        # Save document
        doc.SaveAs(os.path.abspath(full_path))
        word.Visible = True  # Show document after saving
        return f"Document saved successfully at: {full_path}"

    except Exception as e:
        print(f"Save Error: {str(e)}")
        return f"Save Error: {str(e)}"



def type_into_notepad(save_path="C:/VoiceDocs/Notepad"):
    """Opens Notepad, types voice input, and saves automatically."""
    try:
        recognizer = sr.Recognizer()
        
        with sr.Microphone() as source:
            speak("Please speak the text you want to save master")
            audio = recognizer.listen(source, timeout=8)

        text = f"Voice Document\n{datetime.now().strftime('%Y-%m-%d %H:%M')}\n\n{recognizer.recognize_google(audio)}"

        filename = f"VoiceDoc_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        full_path = os.path.join(save_path, filename)

        os.makedirs(save_path, exist_ok=True)

        # Write text to Notepad file
        with open(full_path, 'w') as file:
            file.write(text)

        os.startfile(full_path)  # Open the saved Notepad file
        return f"Document saved successfully at: {full_path}"

    except Exception as e:
        return f"Save Error: {str(e)}"



def open_file(file_path):
    """Opens a file using the default application."""
    try:
        if os.path.exists(file_path):
            subprocess.Popen([file_path], shell=True)
            return f"Opened file: {file_path}"
        else:
            return f"File not found: {file_path} master"
    except Exception as e:
        return f"Error opening file: {str(e)}"

def close_word_files():
    """Close all Microsoft Word files."""
    try:
        subprocess.call(["taskkill", "/f", "/im", "WINWORD.EXE"])
        print("Closed all Word files master.")
    except Exception as e:
        print(f"Error closing Word files: {e}")

def close_notepad_files():
    """Close all Notepad files."""
    try:
        subprocess.call(["taskkill", "/f", "/im", "notepad.exe"])
        print("Closed all Notepad files master.")
    except Exception as e:
        print(f"Error closing Notepad files: {e}")

def close_file(file_path):
    """Closes a file by terminating the process."""
    try:
        if os.path.exists(file_path):
            os.system(f"taskkill /f /im {os.path.basename(file_path)}")
            return f"Closed file: {file_path}"
        else:
            return f"File not found: {file_path}"
    except Exception as e:
        return f"Error closing file: {str(e)}"

def close_google_chrome():
    """Close all Google Chrome browser instances."""
    try:
        os.system("taskkill /f /im chrome.exe")
        print("Closed all Google Chrome instances.")
    except Exception as e:
        print(f"Error closing Google Chrome: {e}")

def extract_keywords(command):
    """Extract keywords from the command."""
    keywords = re.findall(r'\b\w+\b', command)
    for keyword in keywords:
        keyword_frequency[keyword] += 1

def analyze_sentiment(command):
    """Analyze sentiment of the command."""
    analysis = TextBlob(command)
    return analysis.sentiment.polarity  # Returns a value between -1 (negative) and 1 (positive)

def analyze_behavior():
    """Analyze user behavior based on chat history and keyword frequency."""
    if not user_history:
        return "No chat history available master."

    # Most common keyword
    most_common_keyword = max(keyword_frequency, key=keyword_frequency.get, default=None)

    # Analyze command frequency
    frequency_report = "\n".join([f"{cmd}: {len(user_timestamps[cmd])} times" for cmd in user_timestamps])

    # Time analysis (example: last time a command was used)
    time_analysis = {cmd: timestamps[-1].strftime("%Y-%m-%d %H:%M:%S") for cmd, timestamps in user_timestamps.items() if timestamps}

    response = f"Master Your most common keyword is: '{most_common_keyword}'.\n\n" \
               f"Command Frequency:\n{frequency_report}\n\n" \
               f"Last Usage Times:\n" + "\n".join([f"{cmd}: {time}" for cmd, time in time_analysis.items()])

    return response

def open_camera_and_click():
    """Open the camera and take a picture using OpenCV."""
    try:
        # Initialize the camera
        cap = cv2.VideoCapture(0)  # Use 0 for the default camera

        if not cap.isOpened():
            return "Error: Could not open camera."

        # Capture a single frame
        ret, frame = cap.read()
        if ret:
            # Save the captured frame
            save_path = "C:/CapturedImages/"
            os.makedirs(save_path, exist_ok=True)
            filename = f"image_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            full_path = os.path.join(save_path, filename)
            cv2.imwrite(full_path, frame)

            # Release the camera
            cap.release()
            cv2.destroyAllWindows()  # Close any OpenCV windows

            return f"Camera opened and picture taken master. Saved at: {full_path}"
        else:
            cap.release()
            return "Error: Could not capture image."

    except Exception as e:
        return f"Error opening camera: {str(e)}"


def take_screenshot():
    path = "C:/Screenshots"
    os.makedirs(path, exist_ok=True)
    file = os.path.join(path, f"shot_{datetime.now():%Y%m%d_%H%M%S}.png")
    pyautogui.screenshot().save(file)
    return f"Screenshot saved at {file}"

def google_summary(query):
    """Return a clean 3-sentence summary using Google Custom Search snippets."""
    import re
    api_key = 'AIzaSyAo8QiTOZAgQnCa9eaFh9E2bZVsfnj8UQA'  # Replace with your actual API key
    cse_id = 'a203374f897c743d4'  # Replace with your actual CSE ID

    try:
        service = build("customsearch", "v1", developerKey=api_key)
        res = service.cse().list(q=query, cx=cse_id).execute()

        if 'items' in res:
            raw_snippets = [item.get('snippet', '') for item in res['items'][:5]]
            cleaned_sentences = []

            for snippet in raw_snippets:
                # Remove things like "Apr 21, 2021", ellipses, and unwanted characters
                snippet = re.sub(r'\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]* \d{1,2}, \d{4}', '', snippet)
                snippet = snippet.replace("...", "").strip()
                sentences = re.split(r'(?<=[.!?])\s+', snippet)
                cleaned_sentences.extend(sentences)

            # Return the first 3 clean sentences
            summary = " ".join(cleaned_sentences[:3])
            return summary.strip() or "Sorry, I couldn't find a clean summary, master."

        else:
            return "No results found for your query, master."

    except Exception as e:
        return f"Search error: {str(e)}"


@app.route('/')
def index():
    """Render the main page."""
    return render_template('index.html')

@app.route('/ask', methods=['POST'])
def ask():
    user_input = request.json['query'].lower()
    response = ""

    # Log the user input for behavior analysis
    user_history.append(user_input)
    user_behavior[user_input] += 1
    user_timestamps[user_input].append(datetime.now())

    # Extract keywords from the command
    extract_keywords(user_input)

    # Analyze sentiment of the user input
    sentiment_score = analyze_sentiment(user_input)
    sentiment_response = "Your sentiment is neutral."

    if sentiment_score > 0:
        sentiment_response = "Your sentiment is positive."
    elif sentiment_score < 0:
        sentiment_response = "Your sentiment is negative."

    # Command handling
    if 'wikipedia' in user_input:
        user_input = user_input.replace("wikipedia", "").strip()
        results = wikipedia.summary(user_input, sentences=2)
        response = f"According to Wikipedia: {results}"
        speak(response)

    elif 'open camera' in user_input or 'take a photo' in user_input:
        print("Camera command recognized.")  # Debugging output
        response = open_camera_and_click()  # Call the camera function
        speak(response)

    elif 'analyze behavior' in user_input or 'analyze behaviour' in user_input or 'analyse behavior' in user_input or 'analyse behaviour' in user_input:
        response = analyze_behavior()
        speak(response)

    elif 'open youtube' in user_input:
        webbrowser.open("https://youtube.com")
        response = "Opening YouTube."

    elif 'Click a screenshot' in user_input or 'screenshot' in user_input:
        response = take_screenshot()  # Call the screenshot function
        speak(response)

    elif 'open google' in user_input:
        webbrowser.open("https://google.com")
        response = "Opening Google."

    elif 'the time' in user_input:
        strTime = datetime.now().strftime("%H:%M:%S")
        response = f"The time is {strTime}"
        speak(response)

    elif 'search for' in user_input or 'google' in user_input:
        query = user_input.replace("search for", "").replace("google", "").strip()
        response = google_search(query)

    elif 'browse' in user_input or 'search the web for' in user_input:
        search_query = user_input.replace("browse", "").replace("search the web for", "").strip()
        response = browse_topic(search_query)
        speak(response)

    elif 'exit' in user_input or 'quit' in user_input or 'stop' in user_input:
        response = "Goodbye! Have a great day master."
        speak(response)

    elif 'open word' in user_input or 'type my input' in user_input:
        response = type_into_word()
        speak(response)

    elif 'open notepad' in user_input or 'type in notepad' in user_input:
        response = type_into_notepad()
        speak(response)

    elif 'open file' in user_input:
        file_path = user_input.replace("open file", "").strip()
        response = open_file(file_path)
        speak(response)

    elif 'close file' in user_input:
        file_path = user_input.replace("close file", "").strip()
        response = close_file(file_path)
        speak(response)

    elif 'close word' in user_input or 'close all word files' in user_input:
        close_word_files()
        response = "Closed all Word files."
        speak(response)

    elif 'close notepad' in user_input or 'close all notepad files' in user_input:
        close_notepad_files()
        response = "Closed all Notepad files."
        speak(response)

    elif 'close google chrome' in user_input or 'close chrome' in user_input:
        if is_chrome_running():
            close_google_chrome()
            response = "Closed all Google Chrome instances."
        else:
            response = "Google Chrome is not running."
        speak(response)

    elif 'open powerpoint' in user_input or 'Open powerpoint' in user_input or 'Open PowerPoint' in user_input:
        # Specify the path to the PowerPoint shortcut
        power_point_path = r"C:\Users\ADMIN\Desktop\PowerPoint.lnk"
        response = open_application(power_point_path)  # Call with the PowerPoint path
        speak(response)

    elif 'open application' in user_input:
        app_name = user_input.replace("open application", "").strip()
        response = open_application(app_name)
        speak(response)
    
    elif 'refresh' in user_input or 'reload page' in user_input:
        response = "Refreshing the page now, master."
        speak(response)
        return jsonify(response=response, refresh=True)

    else:
        try:
            summary = google_summary(user_input)
            response = summary
            speak(summary)
        except Exception as e:
            response = f"An error occurred: {str(e)}"
            speak(response)

    # Include sentiment response in the final output
    response = f"{sentiment_response}\n{response}"

    return jsonify(response=response, refresh=False)
@app.route('/stop', methods=['POST'])
def stop():
    message = stop_speaking()
    return jsonify(response=message)

if __name__ == "__main__":
    app.run(debug=True)