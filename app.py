import datetime
import speech_recognition as sr
import win32com.client
import pythoncom
import google.generativeai as genai
from PIL import Image
import streamlit as st
import os

# Initialize COM
pythoncom.CoInitialize()

# Initialize Speech
speaker = win32com.client.Dispatch("SAPI.SpVoice")

# Function to say text
def say(text):
    if text[:3] == "AI:":
        text = text[3:]
    print("AI:", text)
    speaker.Speak(f"{text}")

# Function to stop speech
def stop_speech():
    try:
        # Kill the process responsible for speech
        os.system("taskkill /im python.exe /f")  # Forcefully stop Python processes
    except Exception as e:
        st.error(f"Error stopping speech: {e}")

# Function to take voice command
def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"You: {query}")
            return str(query)
        except Exception as e:
            say("Sorry, can you repeat?")
            return "wait"

# API Key Setup (replace with your actual API key)
GENAI_API_KEY = "AIzaSyDBLHAinY4M73akarzljnYKI6nyvhjngpg"
genai.configure(api_key=GENAI_API_KEY)

# Function to generate blog from image with strict word limit
def generate_blog_from_image(image, target_audience, word_limit, mood, tone):
    try:
        prompt = f"""
        Generate a blog based on the content of this image.
        - Target Audience: {target_audience}
        - Word Limit: {word_limit} (strictly {word_limit} words, no more, no less)
        - Mood: {mood}
        - Tone: {tone}
        """
        model = genai.GenerativeModel("gemini-2.0-flash")
        response = model.generate_content([prompt, image])
        blog_text = response.text.strip()

        # Enforce strict word limit
        words = blog_text.split()
        if len(words) > word_limit:
            blog_text = " ".join(words[:word_limit])  # Truncate to the exact word limit
        elif len(words) < word_limit:
            blog_text = blog_text + " ..."  # Add ellipsis if the word count is less than the limit

        return blog_text
    except Exception as e:
        return f"Error generating blog from image: {e}"

# Function to generate blog from text with strict word limit
def generate_blog_from_text(text_input, target_audience, word_limit, mood, tone):
    try:
        prompt = f"""
        Generate a blog based on the following text:
        {text_input}
        - Target Audience: {target_audience}
        - Word Limit: {word_limit} (strictly {word_limit} words, no more, no less)
        - Mood: {mood}
        - Tone: {tone}
        """
        model = genai.GenerativeModel("gemini-2.0-flash")
        response = model.generate_content(prompt)
        blog_text = response.text.strip()

        # Enforce strict word limit
        words = blog_text.split()
        if len(words) > word_limit:
            blog_text = " ".join(words[:word_limit])  # Truncate to the exact word limit
        elif len(words) < word_limit:
            blog_text = blog_text + " ..."  # Add ellipsis if the word count is less than the limit

        return blog_text
    except Exception as e:
        return f"Error generating blog from text: {e}"

# Streamlit UI Code
st.set_page_config(page_title="AI Personal Assistant", page_icon="🤖", layout="wide")

# Custom CSS for enhanced UI
st.markdown("""
    <style>
        .chat-container {
            display: flex;
            flex-direction: column;
            gap: 10px;
            padding: 10px;
            max-height: 70vh;
            overflow-y: auto;
        }

        .user-message {
            align-self: flex-start;
            background-color: #e1ffc7;
            padding: 10px;
            border-radius: 15px;
            max-width: 70%;
            font-size: 16px;
        }

        .ai-message {
            align-self: flex-end;
            background-color: #cfe2f3;
            padding: 10px;
            border-radius: 15px;
            max-width: 70%;
            font-size: 16px;
        }

        .stButton button {
            background-color: #4CAF50;
            color: white;
            border-radius: 5px;
            padding: 10px 20px;
            font-size: 16px;
            transition: 0.3s;
        }

        .stButton button:hover {
            background-color: #45a049;
        }

        .stSelectbox, .stSlider, .stFileUploader {
            background-color: #f0f0f5;
            padding: 10px;
            border-radius: 8px;
        }
    </style>
    """, unsafe_allow_html=True)

# Title and Greeting
st.title("🤖 AI Personal Assistant")

# Check if greeting has already been displayed
if "greeting_displayed" not in st.session_state:
    st.session_state.greeting_displayed = False

# Display greeting only once
if not st.session_state.greeting_displayed:
    current_hour = datetime.datetime.now().hour
    if 5 <= current_hour < 12:
        greeting = "Good morning, Sir!"
    elif 12 <= current_hour < 18:
        greeting = "Good afternoon, Sir!"
    elif 18 <= current_hour < 22:
        greeting = "Good evening, Sir!"
    else:
        greeting = "Hello, Sir!"
    st.write(f"*AI:* {greeting}")
    say(greeting)
    st.session_state.greeting_displayed = True  # Mark greeting as displayed
# Store conversation history in session state
if "conversation" not in st.session_state:
    st.session_state.conversation = []

# Initialize session state for sidebar and input type
if "sidebar_open" not in st.session_state:
    st.session_state.sidebar_open = False

if "input_type" not in st.session_state:
    st.session_state.input_type = "Select an Input Method"

# Initialize session state for read aloud preference
if "read_aloud" not in st.session_state:
    st.session_state.read_aloud = False

# Button to open sidebar
if st.button("Generate a Blog"):
    st.session_state.sidebar_open = not st.session_state.sidebar_open

# Conditional sidebar content based on the button click
if st.session_state.sidebar_open:
    with st.sidebar:
        st.header("AI Assistant Options")
        input_type = st.selectbox("Choose Input Method", ["Select an Input Method", "Speech Input", "Text Input", "Image Input"], key="input_method")
        
        # Update input type in session state
        st.session_state.input_type = input_type

        # Blog customization options
        st.sidebar.header("Blog Customization")
        target_audience = st.selectbox(
            "Target Audience",
            ["General Public", "Students", "Professionals", "Tech Enthusiasts", "Children"]
        )
        word_limit = st.slider("Word Limit", min_value=100, max_value=1000, value=300)
        mood = st.selectbox(
            "Mood",
            ["Informative", "Inspirational", "Funny", "Serious", "Casual"]
        )
        tone = st.selectbox(
            "Tone",
            ["Formal", "Friendly", "Professional", "Humorous", "Neutral"]
        )

        # Read Aloud Toggle
        st.session_state.read_aloud = st.checkbox("Read Blog Aloud Automatically", value=st.session_state.read_aloud)

# Main Content Layout
st.write("### ✨ Let's Generate a Blog!")

# Display input method-specific UI
if st.session_state.input_type == "Speech Input":
    st.write("### 🎤 Speech Input")
    if st.button("Start Listening", key="start_listening_speech"):
        query = takeCommand()
        st.write(f"You: {query}")
        chat = "You: " + query + "\n"
        model = genai.GenerativeModel("gemini-2.0-flash")
        response = model.generate_content(chat)
        st.session_state.conversation.append(f"You: {query}")
        st.session_state.conversation.append(f"AI: {response.text}")
        st.write(f"*AI:* {response.text}")
        if st.session_state.read_aloud:
            say(response.text)

elif st.session_state.input_type == "Text Input":
    st.write("### 📝 Text Input")
    text_input = st.text_area("Enter your topic or information here...", height=150)
    if st.button("Generate Blog from Text", key="generate_blog_from_text"):
        if text_input:
            blog_text = generate_blog_from_text(text_input, target_audience, word_limit, mood, tone)
            st.session_state.conversation.append(f"You: {text_input}")
            st.session_state.conversation.append(f"AI: {blog_text}")
            st.write(f"*AI:* {blog_text}")
            if st.session_state.read_aloud:
                say(blog_text)
        else:
            st.warning("Please enter a topic for the blog.")

elif st.session_state.input_type == "Image Input":
    st.write("### 🖼 Image Input")
    uploaded_file = st.file_uploader("Choose an image...", type=["jpg", "jpeg", "png"])
    if uploaded_file is not None:
        image = Image.open(uploaded_file)
        st.image(image, caption="Uploaded Image", use_column_width=True)
        if st.button("Generate Blog from Image", key="generate_blog_from_image"):
            blog_text = generate_blog_from_image(image, target_audience, word_limit, mood, tone)
            st.session_state.conversation.append(f"AI: {blog_text}")
            st.write(f"*AI:* {blog_text}")
            if st.session_state.read_aloud:
                say(blog_text)

# Display conversation history
st.write("### 💬 Conversation History")
with st.container():
    for message in st.session_state.conversation:
        if message.startswith("You:"):
            st.markdown(f'<div class="user-message">{message[4:]}</div>', unsafe_allow_html=True)
        elif message.startswith("AI:"):
            st.markdown(f'<div class="ai-message">{message[4:]}</div>', unsafe_allow_html=True)

# Manual Read Aloud Button
if st.button("Read Blog Aloud", key="read_blog_aloud"):
    if st.session_state.conversation:
        last_ai_message = st.session_state.conversation[-1]
        if last_ai_message.startswith("AI:"):
            say(last_ai_message[4:])
    else:
        st.warning("No blog generated yet.")

# Stop Reading Button
if st.button("Stop Reading", key="stop_reading"):
    stop_speech()
    st.success("Speech stopped.")

# Provide exit option to stop the assistant
if st.button("Exit Assistant", key="exit_button"):
    say("Goodbye, Sir. Going offline now.")
    st.stop()