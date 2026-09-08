# this is the command to install all the modules in the requirements.txt file
# pip install -r requirements.txt


# venv312\Scripts\activate   --- this line is a comment indicating how to activate the virtual environment in Windows. You can run this command in your command prompt to activate the virtual environment before running the script.
# py -3.12 -m venv venv312  this line is show to how to create a virtual environment using Python 3.12. You can run this command in your command prompt to create a new virtual environment named "venv312".


import os
import datetime
from openai import OpenAI
import wikipedia
import webbrowser
import speech_recognition as sr
import win32com.client as wincl
from groq import Groq
import requests
import json
import re
import shutil
import time
from datetime import datetime
from pathlib import Path
import pyautogui
from PIL import ImageGrab, ImageDraw
from pywinauto import Desktop
from google import genai
import win32gui

from dotenv1 import GROQ_API_KEY
from dotenv1 import NEWS_API_KEY
from dotenv1 import WEATHER_API_KEY
from dotenv1 import OMDB_API_KEY
from dotenv1 import CRICKET_API_KEY
from dotenv1 import OPENROUTER_API_KEY
from dotenv1 import GEMINI_API_KEY



# ---------------------------------------------------------------
# groq ai ki ai 

api = GROQ_API_KEY
conversation_history = [
    {
        "role": "system",
        "content": "You are Jarvis, a helpful voice assistant. The user's name is Pankaj. "
                   "Always remember this and address them as Pankaj when relevant. "
                   "Keep responses concise and natural for voice output."
    }
]

def chat(query):
    client = Groq(api_key=f"{api}")
    global conversation_history
    
    conversation_history.append({"role": "user", "content": query})

    completion = client.chat.completions.create(
        model="qwen/qwen3.6-27b",
        messages=conversation_history,
        temperature=0,
        max_completion_tokens=4096,
        reasoning_effort="none",
        include_reasoning=False,
    )
    
    response_text = completion.choices[0].message.content
    speaker.Speak(response_text)
    
    conversation_history.append({"role": "assistant", "content": response_text})
    return response_text

# -----------------------------------------------------------------------
# ---news ai

# NEWS_API_KEY   # Replace with your actual News API key

def get_news(category="general", country="in"):
    """
    Latest news headlines laata hai
    category: general, business, technology, sports, entertainment, health, science
    country: in (India), us (USA), etc.
    """
    url = f"https://newsapi.org/v2/top-headlines"
    params = {
        "country": country,
        "category": category,
        "apiKey": NEWS_API_KEY,
        "pageSize": 5  # top 5 headlines
    }
    
    try:
        response = requests.get(url, params=params)
        data = response.json()
        
        if data["status"] == "ok" and data["totalResults"] > 0:
            articles = data["articles"]
            news_text = "Here are the top headlines: "
            for i, article in enumerate(articles, 1):
                news_text += f"{i}. {article['title']}. "
            return news_text
        else:
            return "Sorry, I couldn't find any news right now."
    
    except Exception as e:
        print(f"News API error: {e}")
        return "Sorry, I couldn't fetch the news right now."

    # -------------------------------------------------------
# weather ai
# WEATHER_API_KEY 

def get_weather(city="mathura"):
    """
    Diye gaye city ka current weather batata hai
    """
    url = "http://api.weatherapi.com/v1/current.json"
    params = {
        "key": WEATHER_API_KEY,
        "q": city,
        "aqi": "no"  # air quality index nahi chahiye abhi
    }
    
    try:
        response = requests.get(url, params=params)
        data = response.json()
        
        if "error" in data:
            return f"Sorry, I couldn't find weather for {city}."
        
        location = data["location"]["name"]
        temp_c = data["current"]["temp_c"]
        condition = data["current"]["condition"]["text"]
        feels_like = data["current"]["feelslike_c"]
        humidity = data["current"]["humidity"]
        
        weather_text = (f"The weather in {location} is currently {condition}, "
                        f"with a temperature of {temp_c} degrees Celsius, "
                        f"feels like {feels_like} degrees. "
                        f"Humidity is {humidity} percent.")
        
        return weather_text
    
    except Exception as e:
        print(f"Weather API error: {e}")
        return "Sorry, I couldn't fetch the weather right now."
# -------------------------------------------------------
# joke
def get_joke():
    url = "https://official-joke-api.appspot.com/random_joke"
    try:
        response = requests.get(url)
        data = response.json()
        joke = f"{data['setup']} ... {data['punchline']}"
        return joke
    except Exception as e:
        print(f"Joke API error: {e}")
        return "Sorry, I couldn't fetch a joke right now."

# ---------------------------------------------------------------
# motivational quotes
def get_motivational_quote():
    url = "https://api.quotable.io/random"
    params = {"tags": "motivational"}
    try:
        response = requests.get(url, params=params)
        data = response.json()
        quote = f"{data['content']} - {data['author']}"
        return quote
    except Exception as e:
        print(f"Quote API error: {e}")
        return "Stay strong, better days are coming."  # fallback quote

# open cv ke liye
def check_mood_and_motivate(detected_emotion):
    """
    Ye function future me OpenCV se mila emotion input lega
    detected_emotion: "sad", "tired", "stressed" etc.
    """
    if detected_emotion in ["sad", "tired", "stressed"]:
        speaker.Speak(f"You seem {detected_emotion}. Let me share something motivational.")
        quote = get_motivational_quote()
        speaker.Speak(quote)

# ------------------------------------------------------------------------
# dictonary function
def get_word_meaning(word):
    url = f"https://api.dictionaryapi.dev/api/v2/entries/en/{word}"
    try:
        response = requests.get(url)
        data = response.json()
        
        if isinstance(data, list) and len(data) > 0:
            entry = data[0]
            meaning_text = f"{word}: "
            
            for meaning in entry.get("meanings", [])[:1]:  # sirf pehla meaning
                part_of_speech = meaning.get("partOfSpeech", "")
                definitions = meaning.get("definitions", [])
                if definitions:
                    definition = definitions[0].get("definition", "")
                    meaning_text += f"({part_of_speech}) {definition}"
            
            return meaning_text
        else:
            return f"Sorry, I couldn't find the meaning of {word}."
    
    except Exception as e:
        print(f"Dictionary API error: {e}")
        return f"Sorry, I couldn't fetch the meaning of {word}."

# --------------------------------------------------------------------
# movie info function
# OMDB_API_KEY 

def get_movie_info(movie_name):
    url = "http://www.omdbapi.com/"
    params = {
        "apikey": OMDB_API_KEY,
        "t": movie_name
    }
    try:
        response = requests.get(url, params=params)
        data = response.json()
        
        if data.get("Response") == "True":
            title = data.get("Title", "")
            year = data.get("Year", "")
            rating = data.get("imdbRating", "")
            plot = data.get("Plot", "")
            
            movie_text = f"{title} ({year}). IMDb rating: {rating}. Plot: {plot}"
            return movie_text
        else:
            return f"Sorry, I couldn't find information about {movie_name}."
    
    except Exception as e:
        print(f"Movie API error: {e}")
        return "Sorry, I couldn't fetch movie information."

# -----------------------------------------------------------------------------
# cricket
# CRICKET_API_KEY 

def get_live_cricket_scores():
    url = "https://api.cricapi.com/v1/currentMatches"
    params = {"apikey": CRICKET_API_KEY, "offset": 0}
    try:
        response = requests.get(url, params=params)
        data = response.json()
        
        if data.get("status") == "success" and data.get("data"):
            matches = data["data"][:2]  # top 2 matches
            score_text = "Here are the current matches: "
            
            for match in matches:
                name = match.get("name", "")
                status = match.get("status", "")
                score_text += f"{name}. {status}. "
            
            return score_text
        else:
            return "Sorry, no live matches right now."
    
    except Exception as e:
        print(f"Cricket API error: {e}")
        return "Sorry, I couldn't fetch cricket scores."


# ============================================================
#                  OPENROUTER CONFIG
# ============================================================

# OPENROUTER_API_KEY 

client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key=OPENROUTER_API_KEY
)


# ============================================================
#                  main control function
# ============================================================
def control_active_window(query):

    # Gemini vision model
    MODEL = os.getenv(
        "GEMINI_MODEL",
        #"gemini-3.6-flash"
        "gemini-3.1-flash-lite"
        # "gemini-3.1-flash"
        # "gemini-3.5-flash"
        # "gemini-3.5-flash-lite"
    )

    API_KEY = GEMINI_API_KEY

    


    # ============================================================
    # CONFIG
    # ============================================================

    client = genai.Client(api_key=API_KEY)

    SCREENSHOT_DIR = Path("screenshots")
    SCREENSHOT_DIR.mkdir(exist_ok=True)

    MAX_STEPS = 40

    CLICK_VERIFY = True

    WAIT_AFTER_CLICK = 2
    WAIT_AFTER_TYPE = 2
    WAIT_AFTER_ENTER = 2

    API_RETRIES = 2

    pyautogui.PAUSE = 0.08
    pyautogui.FAILSAFE = True


    # ============================================================
    # TASK STATE
    # ============================================================

    task_state = {
        "original_command": "",
        "completed_actions": [],
        "failed_actions": [],
        "rejected_candidates": [],
        "last_action": None,
        "same_action_count": 0,
    }


    # ============================================================
    # SCREENSHOT CLEANUP
    # ============================================================

    def clear_old_screenshots():

        SCREENSHOT_DIR.mkdir(exist_ok=True)

        deleted = 0

        for item in SCREENSHOT_DIR.iterdir():

            try:

                if item.is_file() or item.is_symlink():
                    item.unlink()
                    deleted += 1

                elif item.is_dir():
                    shutil.rmtree(item)
                    deleted += 1

            except Exception as e:

                print(
                    f"[WARN] Could not delete "
                    f"{item}: {e}"
                )

        print(
            f"[SCREENSHOTS] "
            f"Deleted old screenshot items: {deleted}"
        )


    # ============================================================
    # SCREEN
    # ============================================================

    def get_screen_size():

        width, height = pyautogui.size()

        return width, height


    def clamp_coordinates(x, y):

        width, height = get_screen_size()

        x = max(
            0,
            min(int(x), width - 1)
        )

        y = max(
            0,
            min(int(y), height - 1)
        )

        return x, y


    # ============================================================
    # SCREENSHOTS
    # ============================================================

    def take_simple_screenshot(
        step,
        candidate_id=None
    ):

        name = f"step_{step:03d}"

        if candidate_id is not None:
            name += f"_candidate_{candidate_id}"

        path = (
            SCREENSHOT_DIR
            / f"{name}_original.png"
        )

        image = ImageGrab.grab()

        image.save(path)

        return path


    def take_red_circle_screenshot(
        x,
        y,
        step,
        candidate_id=None,
        status="verify"
    ):

        name = f"step_{step:03d}"

        if candidate_id is not None:
            name += f"_candidate_{candidate_id}"

        path = (
            SCREENSHOT_DIR
            / f"{name}_{status}.png"
        )

        image = ImageGrab.grab()

        draw = ImageDraw.Draw(image)

        radius = 22

        # ONLY RED CIRCLE
        # No crosshair
        draw.ellipse(
            (
                x - radius,
                y - radius,
                x + radius,
                y + radius,
            ),
            outline="red",
            width=4
        )

        image.save(path)

        return path


    # ============================================================
    # GEMINI TEXT
    # ============================================================

    def gemini_text(prompt):

        for attempt in range(
            1,
            API_RETRIES + 1
        ):

            try:

                response = client.models.generate_content(
                    model=MODEL,
                    contents=prompt
                )

                if not response:
                    return ""

                if not response.text:
                    return ""

                return response.text.strip()

            except Exception as e:

                error_text = str(e)

                print(
                    f"[GEMINI ERROR] "
                    f"attempt={attempt}: "
                    f"{error_text[:500]}"
                )

                # Don't waste quota on 429
                if (
                    "429" in error_text
                    or
                    "RESOURCE_EXHAUSTED"
                    in error_text
                    or
                    "quota"
                    in error_text.lower()
                ):

                    print(
                        "[GEMINI] "
                        "Quota/rate-limit detected."
                    )

                    return ""

                if attempt < API_RETRIES:

                    time.sleep(1.5)

        return ""


    # ============================================================
    # JSON PARSER
    # ============================================================

    def parse_json(text):

        if not text:
            return None

        text = text.strip()

        # Remove markdown fence
        text = re.sub(
            r"^```json\s*",
            "",
            text,
            flags=re.IGNORECASE
        )

        text = re.sub(
            r"^```\s*",
            "",
            text
        )

        text = re.sub(
            r"\s*```$",
            "",
            text
        )

        try:

            return json.loads(text)

        except Exception:
            pass

        # Find JSON object inside text
        match = re.search(
            r"\{.*\}",
            text,
            flags=re.DOTALL
        )

        if match:

            try:

                return json.loads(
                    match.group(0)
                )

            except Exception:
                pass

        return None


    # ============================================================
    # SAFE UI HELPERS
    # ============================================================

    def safe_name(wrapper):

        try:

            return str(
                wrapper.element_info.name
                or ""
            ).strip()

        except Exception:

            return ""


    def safe_control_type(wrapper):

        try:

            return str(
                wrapper.element_info.control_type
                or ""
            ).strip()

        except Exception:

            return ""


    def safe_visible(wrapper):

        try:

            return bool(
                wrapper.is_visible()
            )

        except Exception:

            return False


    def safe_enabled(wrapper):

        try:

            return bool(
                wrapper.is_enabled()
            )

        except Exception:

            return True


    def safe_rect(wrapper):

        try:

            rect = wrapper.rectangle()

            return {
                "left": int(rect.left),
                "top": int(rect.top),
                "right": int(rect.right),
                "bottom": int(rect.bottom),

                "x": int(
                    (rect.left + rect.right) / 2
                ),

                "y": int(
                    (rect.top + rect.bottom) / 2
                )
            }

        except Exception:

            return None


    def valid_rect(rect):

        if not rect:
            return False

        if (
            rect["right"]
            <= rect["left"]
        ):
            return False

        if (
            rect["bottom"]
            <= rect["top"]
        ):
            return False

        if (
            rect["right"]
            - rect["left"]
            < 2
        ):
            return False

        if (
            rect["bottom"]
            - rect["top"]
            < 2
        ):
            return False

        return True


    # ============================================================
    # CONTROL TYPES
    # ============================================================

    ROOT_TYPES = {
        "Window",
        "Pane",
        "Desktop",
    }

    IGNORE_TYPES = {
        "TitleBar",
        "ScrollBar",
        "Separator",
        "StatusBar",
        "Thumb",
        "Grip",
        "Header",
        "HeaderItem",
        "ToolTip",
    }


    def is_interactive_type(
        control_type
    ):

        return control_type in {

            "Button",
            "Hyperlink",

            "Tab",
            "TabItem",

            "MenuItem",
            "ListItem",
            "TreeItem",

            "CheckBox",
            "RadioButton",

            "ComboBox",

            "SplitButton",

            "Custom",
            "DataItem",
        }


    # ============================================================
    # FOREGROUND WINDOW
    # ============================================================

    def get_foreground_hwnd():

        try:

            hwnd = win32gui.GetForegroundWindow()

            if hwnd:

                return hwnd

        except Exception as e:

            print(
                f"[FOREGROUND ERROR] {e}"
            )

        return None


    def get_active_window_title():

        hwnd = get_foreground_hwnd()

        if not hwnd:
            return ""

        try:

            return (
                win32gui.GetWindowText(hwnd)
                or ""
            ).strip()

        except Exception:

            return ""


    # ============================================================
    # ACTIVE WINDOW SCANNER
    # ============================================================

    def scan_active_window():

        results = []

        hwnd = get_foreground_hwnd()

        if not hwnd:

            return results

        try:

            desktop = Desktop(
                backend="uia"
            )

            # IMPORTANT:
            # Foreground HWND se direct window lo.
            active = desktop.window(
                handle=hwnd
            )

            if not active.exists():

                return results

        except Exception as e:

            print(
                f"[ACTIVE WINDOW ERROR] "
                f"{e}"
            )

            return results

        try:

            # IMPORTANT:
            # descendants() mein visible_only mat do.
            descendants = active.descendants()

        except Exception as e:

            print(
                f"[DESCENDANTS ERROR] "
                f"{e}"
            )

            return results

        next_id = 1

        screen_w, screen_h = (
            get_screen_size()
        )

        for element in descendants:

            try:

                name = safe_name(element)

                control_type = (
                    safe_control_type(element)
                )

                if not name:
                    continue

                if not safe_visible(element):
                    continue

                if not safe_enabled(element):
                    continue

                if control_type in IGNORE_TYPES:
                    continue

                # Root panes/windows ko click candidate
                # nahi banana hai.
                if control_type in ROOT_TYPES:
                    continue

                rect = safe_rect(element)

                if not valid_rect(rect):
                    continue

                if (
                    rect["x"] < 0
                    or rect["x"] >= screen_w
                ):
                    continue

                if (
                    rect["y"] < 0
                    or rect["y"] >= screen_h
                ):
                    continue

                results.append({

                    "id": next_id,

                    "name": name,

                    "type": control_type,

                    "source": "ACTIVE_WINDOW",

                    "x": rect["x"],

                    "y": rect["y"],

                    "left": rect["left"],

                    "top": rect["top"],

                    "right": rect["right"],

                    "bottom": rect["bottom"],

                    "clickable":
                        is_interactive_type(
                            control_type
                        ),

                    "wrapper": element,
                })

                next_id += 1

            except Exception:

                continue

        return results


    # ============================================================
    # TASKBAR SCANNER
    # ============================================================

    def scan_taskbar():

        results = []

        try:

            desktop = Desktop(
                backend="uia"
            )

            taskbar = desktop.window(
                class_name="Shell_TrayWnd"
            )

            if not taskbar.exists():

                return results

        except Exception as e:

            print(
                f"[TASKBAR WINDOW ERROR] "
                f"{e}"
            )

            return results

        try:

            # IMPORTANT:
            # visible_only=True hata diya.
            descendants = (
                taskbar.descendants()
            )

        except Exception as e:

            print(
                f"[TASKBAR DESCENDANTS ERROR] "
                f"{e}"
            )

            return results

        next_id = 100000

        for element in descendants:

            try:

                name = safe_name(element)

                control_type = (
                    safe_control_type(element)
                )

                if not name:
                    continue

                if not safe_visible(element):
                    continue

                if not safe_enabled(element):
                    continue

                if control_type in IGNORE_TYPES:
                    continue

                if not is_interactive_type(
                    control_type
                ):
                    continue

                rect = safe_rect(element)

                if not valid_rect(rect):
                    continue

                results.append({

                    "id": next_id,

                    "name": name,

                    "type": control_type,

                    "source": "TASKBAR",

                    "x": rect["x"],

                    "y": rect["y"],

                    "left": rect["left"],

                    "top": rect["top"],

                    "right": rect["right"],

                    "bottom": rect["bottom"],

                    "clickable": True,

                    "wrapper": element,
                })

                next_id += 1

            except Exception:

                continue

        return results


    # ============================================================
    # UI DEDUPLICATION
    # ============================================================

    def deduplicate_ui(
        elements
    ):

        result = []

        seen = set()

        for item in elements:

            key = (
                item["id"],
                item["name"],
                item["type"],
                item["source"]
            )

            if key in seen:
                continue

            seen.add(key)

            result.append(item)

        return result


    # ============================================================
    # UI CONTEXT
    # ============================================================

    def build_ui_context(
        elements,
        max_items=200
    ):

        if not elements:

            return (
                "[NO UI ELEMENTS DETECTED]\n"
                "Do NOT invent a target_id."
            )

        lines = []

        for item in elements[:max_items]:

            name = item["name"]

            name = name.replace(
                "\n",
                " "
            )

            name = re.sub(
                r"\s+",
                " ",
                name
            )

            if len(name) > 120:

                name = (
                    name[:117]
                    + "..."
                )

            lines.append(
                f'id={item["id"]} | '
                f'name="{name}" | '
                f'type={item["type"]} | '
                f'source={item["source"]}'
            )

        return "\n".join(lines)


    # ============================================================
    # TASK HISTORY
    # ============================================================

    def compact_history():

        return {

            "original_command":
                task_state[
                    "original_command"
                ],

            "completed_actions":
                task_state[
                    "completed_actions"
                ][-25:],

            "failed_actions":
                task_state[
                    "failed_actions"
                ][-15:],

            "rejected_candidates":
                task_state[
                    "rejected_candidates"
                ][-20:],

            "last_action":
                task_state[
                    "last_action"
                ],

            "same_action_count":
                task_state[
                    "same_action_count"
                ],
        }


    # ============================================================
    # ACTION SIGNATURE
    # ============================================================

    def action_signature(
        action
    ):

        if not action:
            return ""

        action_type = str(
            action.get(
                "action",
                ""
            )
        ).upper()

        target_id = action.get(
            "target_id",
            ""
        )

        target_name = action.get(
            "target_name",
            ""
        )

        text = action.get(
            "text",
            ""
        )

        key = action.get(
            "key",
            ""
        )

        return (
            f"{action_type}|"
            f"{target_id}|"
            f"{target_name}|"
            f"{text}|"
            f"{key}"
        )


    # ============================================================
    # GEMINI PLANNER
    # ============================================================

    def ask_gemini_next_action(
        command,
        active_window,
        ui_elements
    ):

        ui_context = build_ui_context(
            ui_elements
        )

        history = compact_history()

        prompt = f"""
    You are the planning brain of a Windows computer-control agent.

    Your job is to COMPLETE the user's ORIGINAL COMMAND.

    ha agar user kuch video song play karne ki baat kar rha he chrome par to youtube kholna he 

    There is NO predefined workflow.

    Do NOT assume:
    - Chrome always comes first
    - YouTube always comes next
    - LinkedIn always comes next
    - Search is always the same
    - Every command has the same number of steps

    Understand the actual natural-language command.

    --------------------------------------------------
    ORIGINAL COMMAND
    --------------------------------------------------

    {command}

    --------------------------------------------------
    CURRENT ACTIVE WINDOW
    --------------------------------------------------

    {active_window}

    --------------------------------------------------
    TASK HISTORY
    --------------------------------------------------

    {json.dumps(
        history,
        indent=2,
        ensure_ascii=False
    )}

    --------------------------------------------------
    CURRENT UI ELEMENTS
    --------------------------------------------------

    {ui_context}

    --------------------------------------------------
    IMPORTANT RULES
    --------------------------------------------------

    

    1. ORIGINAL COMMAND is the source of truth.

    2. COMPLETED ACTIONS tell you what has already happened.

    3. FAILED ACTIONS tell you what did NOT work.

    4. Never declare DONE just because one part of a
       multi-part command is finished.

    5. Only return DONE when the complete original command
       has been fulfilled.

    6. If the browser/app is already open but the requested
       website/page is not open, continue navigation.

    7. If the requested page is already open, use it.

    8. If a target exists in CURRENT UI ELEMENTS,
       CLICK must use its exact target_id.

    9. NEVER invent target IDs.

    10. If CURRENT UI ELEMENTS says:
           [NO UI ELEMENTS DETECTED]
        then do NOT use CLICK with an invented ID.

    11. When UIA does not expose a control but a keyboard
        shortcut can perform the next action, use KEY.

    12. KEY can be used for generic keyboard actions.

    13. Examples of valid KEY values:
           WIN
           ENTER
           ESC
           TAB
           CTRL+L
           CTRL+A
           ALT+TAB
           BACKSPACE
           CTRL+C
           CTRL+V

    14. Do not use a keyboard shortcut just because it is
        convenient if a visible UI target is clearly available.

    15. TYPE should normally target a real UI element.

    16. If the current UI has an address bar, search box,
        textbox, etc., use its exact target_id.

    17. If an application needs to be opened and there is
        no visible UI element for it, KEY can be used to open
        the Windows Start/search interface.

    18. Interpret obvious spelling mistakes from context.

    19. Do not perform unrelated actions.

    20. Execute ONE action at a time.

    21. After an action, the program will scan the UI again
        and ask you for the next action.

    22. If a previous action failed, choose a genuinely
        different approach instead of blindly repeating it.

    23. If the command asks for multiple things joined by
        "and", "then", "after that", etc., complete them all.

    24. If a website/search is requested, do not stop before
        the requested navigation/search has actually happened.

    25. For a media task, only open/play a result when the
        original command actually requires opening/playing it.

    26. when youtube video is paused then click on play button to play the video.

    --------------------------------------------------
    ALLOWED OUTPUTS
    --------------------------------------------------

    CLICK:

    {{
      "action": "CLICK",
      "target_id": 123,
      "reason": "short reason"
    }}

    TYPE:

    {{
      "action": "TYPE",
      "target_id": 123,
      "text": "exact text",
      "reason": "short reason"
    }}

    KEY:

    {{
      "action": "KEY",
      "key": "CTRL+L",
      "reason": "short reason"
    }}

    ENTER:

    {{
      "action": "ENTER",
      "reason": "short reason"
    }}

    WAIT:

    {{
      "action": "WAIT",
      "seconds": 2,
      "reason": "short reason"
    }}

    DONE:

    {{
      "action": "DONE",
      "reason": "explain why the ORIGINAL COMMAND is fully complete"
    }}

    --------------------------------------------------
    FINAL RULE
    --------------------------------------------------

    Return ONLY valid JSON.

    Do not return markdown.

    Do not return explanations outside JSON.
    """

        raw = gemini_text(
            prompt
        )

        print("\n[GEMINI RAW]")
        print(raw)

        result = parse_json(
            raw
        )

        if not isinstance(
            result,
            dict
        ):

            print(
                "[PLANNER] Invalid JSON."
            )

            return None

        result["action"] = str(
            result.get(
                "action",
                ""
            )
        ).upper().strip()

        return result


    # ============================================================
    # FIND ELEMENT
    # ============================================================

    def find_element_by_id(
        elements,
        target_id
    ):

        try:

            target_id = int(
                target_id
            )

        except Exception:

            return None

        for item in elements:

            if item["id"] == target_id:

                return item

        return None


    # ============================================================
    # FIND BY NAME
    # ============================================================

    def find_same_name_candidates(
        elements,
        target_name
    ):

        target_name = str(
            target_name or ""
        ).strip()

        result = []

        for item in elements:

            if (
                item["name"]
                .strip()
                .lower()
                ==
                target_name.lower()
            ):

                result.append(item)

        return result


    # ============================================================
    # VISUAL VERIFICATION
    # ============================================================

    def verify_click_target(
        original_path,
        marked_path,
        candidate
    ):

        prompt = f"""
    You are verifying a Windows UI click target.

    IMAGE 1:
    Original screenshot.

    IMAGE 2:
    Same screenshot with ONE red circle marking the
    exact location that will be clicked.

    Candidate:
    Name: {candidate["name"]}
    Type: {candidate["type"]}
    Source: {candidate["source"]}

    Determine whether the red circle is actually over
    the visible UI object represented by the candidate.

    Return ONLY:

    {{
      "result": "CORRECT",
      "reason": "short explanation"
    }}

    OR

    {{
      "result": "WRONG",
      "reason": "short explanation"
    }}

    Rules:
    - when the application is open in taskbar allready so click on it not to search in start menu.
    - 
    - CORRECT only when the marked point clearly belongs
      to the candidate.
    - If uncertain, return WRONG.
    - Do not guess.
    """

        try:

            with open(
                original_path,
                "rb"
            ) as f:

                original_bytes = f.read()

            with open(
                marked_path,
                "rb"
            ) as f:

                marked_bytes = f.read()

            contents = [

                prompt,

                {
                    "inline_data": {
                        "mime_type":
                            "image/png",

                        "data":
                            original_bytes
                    }
                },

                {
                    "inline_data": {
                        "mime_type":
                            "image/png",

                        "data":
                            marked_bytes
                    }
                }
            ]

            response = (
                client.models.generate_content(
                    model=MODEL,
                    contents=contents
                )
            )

            raw = (
                response.text.strip()
                if response
                and response.text
                else ""
            )

            print(
                "\n[VISUAL VERIFY RAW]"
            )

            print(raw)

            result = parse_json(
                raw
            )

            if not isinstance(
                result,
                dict
            ):

                return (
                    False,
                    "Invalid verification response"
                )

            answer = str(
                result.get(
                    "result",
                    ""
                )
            ).upper().strip()

            reason = str(
                result.get(
                    "reason",
                    ""
                )
            ).strip()

            return (
                answer == "CORRECT",
                reason
            )

        except Exception as e:

            print(
                f"[VISUAL VERIFY ERROR] "
                f"{str(e)[:500]}"
            )

            return (
                False,
                str(e)
            )


    # ============================================================
    # CLICK CANDIDATE
    # ============================================================

    def click_candidate(
        candidate,
        step
    ):

        x, y = clamp_coordinates(
            candidate["x"],
            candidate["y"]
        )

        print(
            f"[CLICK] "
            f'id={candidate["id"]} '
            f'name="{candidate["name"]}" '
            f'type={candidate["type"]} '
            f'position=({x},{y})'
        )

        original = (
            take_simple_screenshot(
                step,
                candidate["id"]
            )
        )

        marked = (
            take_red_circle_screenshot(
                x,
                y,
                step,
                candidate["id"],
                "verify"
            )
        )

        if CLICK_VERIFY:

            correct, reason = (
                verify_click_target(
                    original,
                    marked,
                    candidate
                )
            )

            print(
                f"[VERIFY] "
                f"{'CORRECT' if correct else 'WRONG'} "
                f"- {reason}"
            )

            if not correct:

                return False

        pyautogui.click(
            x,
            y
        )

        time.sleep(
            WAIT_AFTER_CLICK
        )

        return True


    # ============================================================
    # TYPE
    # ============================================================

    def execute_type(
        action,
        elements
    ):

        target_id = action.get(
            "target_id"
        )

        text = action.get(
            "text"
        )

        if text is None:

            return (
                False,
                "TYPE text missing"
            )

        candidate = None

        if target_id is not None:

            candidate = (
                find_element_by_id(
                    elements,
                    target_id
                )
            )

        if candidate is None:

            target_name = action.get(
                "target_name"
            )

            if target_name:

                candidates = (
                    find_same_name_candidates(
                        elements,
                        target_name
                    )
                )

                if candidates:

                    candidate = (
                        candidates[0]
                    )

        if candidate is None:

            return (
                False,
                "TYPE target not found"
            )

        x, y = clamp_coordinates(
            candidate["x"],
            candidate["y"]
        )

        print(
            f'[TYPE] "{text}" '
            f'into "{candidate["name"]}"'
        )

        pyautogui.click(
            x,
            y
        )

        time.sleep(
            0.25
        )

        pyautogui.hotkey(
            "ctrl",
            "a"
        )

        time.sleep(
            0.1
        )

        # write() is safer for normal English text.
        pyautogui.write(
            str(text),
            interval=0.01
        )

        time.sleep(
            WAIT_AFTER_TYPE
        )

        return (
            True,
            "Text typed successfully"
        )


    # ============================================================
    # KEY ACTION
    # ============================================================

    def execute_key(
        action
    ):

        key_text = str(
            action.get(
                "key",
                ""
            )
        ).strip().upper()

        if not key_text:

            return (
                False,
                "KEY missing"
            )

        print(
            f"[KEY] {key_text}"
        )

        # --------------------------------------------------------
        # Single keys
        # --------------------------------------------------------

        single_keys = {
            "ENTER": "enter",
            "ESC": "esc",
            "ESCAPE": "esc",
            "TAB": "tab",
            "SPACE": "space",
            "BACKSPACE": "backspace",
            "DELETE": "delete",
            "UP": "up",
            "DOWN": "down",
            "LEFT": "left",
            "RIGHT": "right",
            "HOME": "home",
            "END": "end",
            "PAGEUP": "pageup",
            "PAGEDOWN": "pagedown",
            "WIN": "win",
            "WINDOWS": "win",
            "F1": "f1",
            "F2": "f2",
            "F3": "f3",
            "F4": "f4",
            "F5": "f5",
            "F6": "f6",
            "F7": "f7",
            "F8": "f8",
            "F9": "f9",
            "F10": "f10",
            "F11": "f11",
            "F12": "f12",
        }

        if key_text in single_keys:

            pyautogui.press(
                single_keys[key_text]
            )

        else:

            # Example:
            # CTRL+L
            # ALT+TAB
            # CTRL+SHIFT+ESC

            parts = [
                p.strip().lower()
                for p in key_text.split("+")
                if p.strip()
            ]

            if not parts:

                return (
                    False,
                    "Invalid KEY"
                )

            # Map Windows -> win
            mapped = []

            for key in parts:

                if key in {
                    "windows",
                    "window",
                    "win"
                }:

                    mapped.append("win")

                elif key in {
                    "control",
                    "ctrl"
                }:

                    mapped.append("ctrl")

                elif key in {
                    "alt"
                }:

                    mapped.append("alt")

                elif key in {
                    "shift"
                }:

                    mapped.append("shift")

                elif key in {
                    "escape"
                }:

                    mapped.append("esc")

                else:

                    mapped.append(key)

            pyautogui.hotkey(
                *mapped
            )

        time.sleep(
            WAIT_AFTER_CLICK
        )

        return (
            True,
            f"Key executed: {key_text}"
        )


    # ============================================================
    # ENTER
    # ============================================================

    def execute_enter():

        print(
            "[ENTER]"
        )

        pyautogui.press(
            "enter"
        )

        time.sleep(
            WAIT_AFTER_ENTER
        )

        return (
            True,
            "Enter pressed"
        )


    # ============================================================
    # WAIT
    # ============================================================

    def execute_wait(
        action
    ):

        try:

            seconds = float(
                action.get(
                    "seconds",
                    1.5
                )
            )

        except Exception:

            seconds = 1.5

        seconds = max(
            0.2,
            min(seconds, 10)
        )

        print(
            f"[WAIT] {seconds} seconds"
        )

        time.sleep(
            seconds
        )

        return (
            True,
            "Wait completed"
        )


    # ============================================================
    # VALIDATE ACTION
    # ============================================================

    ALLOWED_ACTIONS = {
        "CLICK",
        "TYPE",
        "KEY",
        "ENTER",
        "WAIT",
        "DONE",
    }


    def validate_action(
        action
    ):

        if not isinstance(
            action,
            dict
        ):

            return (
                False,
                "Action is not object"
            )

        action_type = str(
            action.get(
                "action",
                ""
            )
        ).upper().strip()

        if action_type not in ALLOWED_ACTIONS:

            return (
                False,
                f"Unsupported action: {action_type}"
            )

        if action_type == "CLICK":

            if (
                action.get("target_id")
                is None
                and
                not action.get(
                    "target_name"
                )
            ):

                return (
                    False,
                    "CLICK needs target_id"
                )

        if action_type == "TYPE":

            if action.get("text") is None:

                return (
                    False,
                    "TYPE text missing"
                )

        if action_type == "KEY":

            if not action.get("key"):

                return (
                    False,
                    "KEY missing"
                )

        return (
            True,
            ""
        )


    # ============================================================
    # EXECUTE CLICK
    # ============================================================

    def execute_click(
        action,
        elements,
        step
    ):

        candidates = []

        target_id = action.get(
            "target_id"
        )

        if target_id is not None:

            candidate = (
                find_element_by_id(
                    elements,
                    target_id
                )
            )

            if candidate:

                candidates.append(
                    candidate
                )

        # Fallback to exact name
        if not candidates:

            target_name = action.get(
                "target_name"
            )

            if target_name:

                candidates = (
                    find_same_name_candidates(
                        elements,
                        target_name
                    )
                )

        if not candidates:

            return (
                False,
                "Click target not found"
            )

        # Remove rejected candidates
        usable = []

        for candidate in candidates:

            if candidate["id"] in (
                task_state[
                    "rejected_candidates"
                ]
            ):

                continue

            usable.append(
                candidate
            )

        if not usable:

            return (
                False,
                "All matching candidates rejected"
            )

        # Try candidates one by one
        for candidate in usable:

            ok = click_candidate(
                candidate,
                step
            )

            if ok:

                return (
                    True,
                    f'Clicked "{candidate["name"]}"'
                )

            task_state[
                "rejected_candidates"
            ].append(
                candidate["id"]
            )

        return (
            False,
            "All candidates failed verification"
        )


    # ============================================================
    # HISTORY
    # ============================================================

    def add_completed_action(
        action,
        result
    ):

        task_state[
            "completed_actions"
        ].append({

            "action": action,

            "result": result,

            "time":
                time.strftime(
                    "%H:%M:%S"
                )
        })


    def add_failed_action(
        action,
        reason
    ):

        task_state[
            "failed_actions"
        ].append({

            "action": action,

            "reason": reason,

            "time":
                time.strftime(
                    "%H:%M:%S"
                )
        })


    # ============================================================
    # MAIN AGENT
    # ============================================================

    def run_agent(
        command
    ):

        global task_state

        command = str(
            command
        ).strip()

        if not command:

            print(
                "[ERROR] Empty command."
            )

            return

        # --------------------------------------------------------
        # NEW TASK
        # --------------------------------------------------------

        clear_old_screenshots()

        task_state = {

            "original_command":
                command,

            "completed_actions":
                [],

            "failed_actions":
                [],

            "rejected_candidates":
                [],

            "last_action":
                None,

            "same_action_count":
                0,
        }

        print()
        print("=" * 70)
        print("NEW TASK")
        print("=" * 70)
        print(command)
        print("=" * 70)

        # --------------------------------------------------------
        # LOOP
        # --------------------------------------------------------

        for step in range(
            1,
            MAX_STEPS + 1
        ):

            print()
            print("-" * 70)
            print(
                f"STEP {step}"
            )
            print("-" * 70)

            # ----------------------------------------------------
            # ACTIVE WINDOW
            # ----------------------------------------------------

            active_window = (
                get_active_window_title()
            )

            print(
                f"[ACTIVE WINDOW] "
                f"{active_window}"
            )

            # ----------------------------------------------------
            # SCAN
            # ----------------------------------------------------

            active_elements = (
                scan_active_window()
            )

            taskbar_elements = (
                scan_taskbar()
            )

            elements = deduplicate_ui(
                active_elements
                +
                taskbar_elements
            )

            print(
                f"[UI] "
                f"Active={len(active_elements)} "
                f"Taskbar={len(taskbar_elements)} "
                f"Total={len(elements)}"
            )

            # ----------------------------------------------------
            # PRINT UI
            # ----------------------------------------------------

            for item in elements[:80]:

                print(
                    f'  ID={item["id"]} | '
                    f'{item["type"]} | '
                    f'{item["source"]} | '
                    f'"{item["name"]}"'
                )

            # ----------------------------------------------------
            # GEMINI
            # ----------------------------------------------------

            action = (
                ask_gemini_next_action(
                    command,
                    active_window,
                    elements
                )
            )

            if not action:

                print(
                    "[STOP] "
                    "Gemini planner failed."
                )

                return

            # ----------------------------------------------------
            # VALIDATE
            # ----------------------------------------------------

            valid, reason = (
                validate_action(
                    action
                )
            )

            if not valid:

                print(
                    f"[INVALID ACTION] "
                    f"{reason}"
                )

                add_failed_action(
                    action,
                    reason
                )

                continue

            action_type = action[
                "action"
            ]

            print()
            print(
                f"[NEXT ACTION] "
                f"{action_type}"
            )

            print(
                json.dumps(
                    action,
                    indent=2,
                    ensure_ascii=False
                )
            )

            # ----------------------------------------------------
            # SAME ACTION PROTECTION
            # ----------------------------------------------------

            signature = (
                action_signature(
                    action
                )
            )

            last_signature = (
                action_signature(
                    task_state[
                        "last_action"
                    ]
                )
            )

            if (
                signature
                and
                signature == last_signature
            ):

                task_state[
                    "same_action_count"
                ] += 1

            else:

                task_state[
                    "same_action_count"
                ] = 0

            task_state[
                "last_action"
            ] = action

            if (
                task_state[
                    "same_action_count"
                ] >= 3
            ):

                print(
                    "[STOP] "
                    "Same action repeated 3 times."
                )

                add_failed_action(
                    action,
                    "Repeated action loop"
                )

                return

            # ----------------------------------------------------
            # DONE
            # ----------------------------------------------------

            if action_type == "DONE":

                print()
                print("=" * 70)
                print("TASK COMPLETED")
                print("=" * 70)

                reason = str(
                    action.get(
                        "reason",
                        ""
                    )
                ).strip()

                if reason:

                    print(
                        f"Reason: {reason}"
                    )

                print("=" * 70)

                return

            # ----------------------------------------------------
            # CLICK
            # ----------------------------------------------------

            if action_type == "CLICK":

                success, result = (
                    execute_click(
                        action,
                        elements,
                        step
                    )
                )

                if success:

                    print(
                        f"[SUCCESS] "
                        f"{result}"
                    )

                    add_completed_action(
                        action,
                        result
                    )

                    task_state[
                        "same_action_count"
                    ] = 0

                else:

                    print(
                        f"[CLICK FAILED] "
                        f"{result}"
                    )

                    add_failed_action(
                        action,
                        result
                    )

                continue

            # ----------------------------------------------------
            # TYPE
            # ----------------------------------------------------

            if action_type == "TYPE":

                success, result = (
                    execute_type(
                        action,
                        elements
                    )
                )

                if success:

                    print(
                        f"[SUCCESS] "
                        f"{result}"
                    )

                    add_completed_action(
                        action,
                        result
                    )

                    task_state[
                        "same_action_count"
                    ] = 0

                else:

                    print(
                        f"[TYPE FAILED] "
                        f"{result}"
                    )

                    add_failed_action(
                        action,
                        result
                    )

                continue

            # ----------------------------------------------------
            # KEY
            # ----------------------------------------------------

            if action_type == "KEY":

                success, result = (
                    execute_key(
                        action
                    )
                )

                if success:

                    add_completed_action(
                        action,
                        result
                    )

                    task_state[
                        "same_action_count"
                    ] = 0

                else:

                    print(
                        f"[KEY FAILED] "
                        f"{result}"
                    )

                    add_failed_action(
                        action,
                        result
                    )

                continue

            # ----------------------------------------------------
            # ENTER
            # ----------------------------------------------------

            if action_type == "ENTER":

                success, result = (
                    execute_enter()
                )

                if success:

                    add_completed_action(
                        action,
                        result
                    )

                    task_state[
                        "same_action_count"
                    ] = 0

                else:

                    add_failed_action(
                        action,
                        result
                    )

                continue

            # ----------------------------------------------------
            # WAIT
            # ----------------------------------------------------

            if action_type == "WAIT":

                success, result = (
                    execute_wait(
                        action
                    )
                )

                if success:

                    add_completed_action(
                        action,
                        result
                    )

                    task_state[
                        "same_action_count"
                    ] = 0

                else:

                    add_failed_action(
                        action,
                        result
                    )

                continue

        # --------------------------------------------------------
        # MAX STEPS
        # --------------------------------------------------------

        print()
        print("=" * 70)
        print("TASK STOPPED")
        print("=" * 70)

        print(
            f"Maximum steps "
            f"({MAX_STEPS}) reached."
        )

        print("=" * 70)


    # ============================================================
    # PROGRAM
    # ============================================================

    if __name__ == "__main__":

        print("=" * 70)
        print(
            "GENERIC GEMINI WINDOWS UI AGENT"
        )
        print("=" * 70)

        print(
            f"Model: {MODEL}"
        )

        print(
            "Type 'exit' to close."
        )

        while True:

            try:

                command = query.strip()

            except KeyboardInterrupt:

                print(
                    "\nExiting..."
                )

                break

            except EOFError:

                print(
                    "\nExiting..."
                )

                break

            if not command:
                continue

            if command.lower() in {
                "exit",
                "quit"
            }:

                print(
                    "Exiting..."
                )

                break

            try:

                run_agent(
                    command
                )

            except KeyboardInterrupt:

                print(
                    "\n[STOPPED BY USER]"
                )

            except Exception as e:

                print(
                    f"\n[FATAL ERROR] "
                    f"{type(e).__name__}: {e}"
                )

# -----------------    # 
# command ko lene ka kaam karta he 

speaker = wincl.Dispatch("SAPI.SpVoice")

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        r.adjust_for_ambient_noise(source, duration=5)
        audio = r.listen(source, timeout=None)
        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except sr.RequestError as e:
            speaker.Speak("Could not request results; {0}".format(e))
            return ""
        except Exception as e:
            # speaker.Speak("Sorry, I did not understand. Please try again.")
            return ""
# -------------------------------------

# main code     -------------------------------------------------------------------------------------------------------------------

if __name__ == "__main__":
    print("hello")
    speaker.Speak(" hello, I am your AI assistant. How can I help you today?")
    # ---------------------------------------------------
    while True:
        print("Listening...")
        query = takeCommand()
        # speaker.Speak(text)
        # ---------------------------------------------------
        # open sites and browsers
        site = [
            ("Chrome", "google.com"),
            ("Google", "google.com"),
            ("YouTube", "youtube.com"),
            ("Facebook", "facebook.com"),
            ("Twitter", "twitter.com"),
            ("Instagram", "instagram.com"),
            ("LinkedIn", "linkedin.com"),
            ("Reddit", "reddit.com"),
            ("Wikipedia", "wikipedia.org"),
            ("Amazon", "amazon.com"),
            ("Netflix", "netflix.com")
        ]
        for sites in site:
            if f"Open {sites[0]}".lower() in query.lower():
                speaker.Speak(f"Opening {sites[0]}")
                webbrowser.open(f"https://www.{sites[1]}")
                site_opened = True
                break
        # if site_opened:
        #     continue
        # ---------------------------------------------------
        # songs
        if "open music".lower() in query.lower():
            music_dir = "C:\\Users\\YourUsername\\Music"  # Change this to your music directory
            # os.system(f"open {music_dir}")
            songs = os.listdir(music_dir)
            if songs:
                os.startfile(os.path.join(music_dir, songs[0]))
                speaker.Speak("Playing music")
            else:
                speaker.Speak("No music files found in the directory.")
        # ---------------------------------------------------
        # time
        elif "the time".lower() in query.lower():
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speaker.Speak(f"The time is {strTime}")
        # ---------------------------------------------------
        # wikipedia
        elif "wikipedia".lower() in query.lower():
            speaker.Speak("Searching Wikipedia...")
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speaker.Speak("According to Wikipedia")
            print(results)
            speaker.Speak(results)
        # ------------------------------------------------------
        # news

        elif "news" in query.lower() or "headlines" in query.lower():
            speaker.Speak("Fetching the latest news for you")

            # Category detect karo command se (optional)
            if "sports" in query.lower():
                news = get_news(category="sports")
            elif "technology" in query.lower() or "tech" in query.lower():
                news = get_news(category="technology")
            elif "business" in query.lower():
                news = get_news(category="business")
            else:
                news = get_news(category="general")

            print(news)
            speaker.Speak(news)
        # ------------------------------------------------------
        # weather
        elif "weather" in query.lower():
            speaker.Speak("Fetching the weather for you")

            # City extract karo command se (agar bola gaya ho)
            if " in " in query.lower():
                city = query.lower().split(" in ")[-1].strip()
                weather = get_weather(city=city)
            else:
                weather = get_weather(city="mathura")  # default city, apna shehar daal do

            print(weather)
            speaker.Speak(weather)

        # ------------------------------------------------------
        # jokes
        elif "joke" in query.lower():
            joke = get_joke()
            print(joke)
            speaker.Speak(joke)

        # ---------------------------------------------------
        # motivational quotes
        elif "motivational" in query.lower() or "motivate me" in query.lower():
            quote = get_motivational_quote()
            print(quote)
            speaker.Speak(quote)

        # ---------------------------------------------------
        # dictonary
        elif "meaning of" in query.lower():
            word = query.lower().split("meaning of")[-1].strip()
            meaning = get_word_meaning(word)
            print(meaning)
            speaker.Speak(meaning)

        # ---------------------------------------------------
        # search movies
        elif "tell me about the movie" in query.lower() or "movie info" in query.lower():
            movie_name = query.lower().replace("tell me about the movie", "").replace("movie info", "").strip()
            movie_info = get_movie_info(movie_name)
            print(movie_info)
            speaker.Speak(movie_info)

        # ---------------------------------------------------
        elif "cricket score" in query.lower() or "cricket match" in query.lower():
            scores = get_live_cricket_scores()
            print(scores)
            speaker.Speak(scores)

        # ---------------------------------------------------
        # controlling

        if "full control".lower() in query.lower() or "hello ai".lower() in query.lower():
            speaker.Speak("Hello! sir, give me your commands,")
            while True:
                print("Listening...")
                query = takeCommand()
                if query == "":
                    continue

                control_active_window(query)

                if "exit ai".lower() in query.lower() or "quit ai".lower() in query.lower() or "exit nik".lower() in query.lower():
                    speaker.Speak("Exiting the AI conversation.")
                    break

        # ---------------------------------------------------
        # chatting with ai
        elif query != "":
            print("chatting")
            chat(query)
