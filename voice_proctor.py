import os
import re
import sys
import json
import time
import queue
from xml.sax.saxutils import escape

import winsound
import pyautogui
import sounddevice as sd
from vosk import Model, KaldiRecognizer, SetLogLevel

try:
    from vosk import EndpointerMode
except ImportError:
    EndpointerMode = None

try:
    import win32com.client
except ImportError:
    print("Missing package. Run: pip install pywin32")
    sys.exit(1)


MODEL_PATH = "model"
VOSK_LOG_LEVEL = -1            

COURSE_SPOKEN = "C S C two zero one"          
COURSE_TITLE = "Introduction to Computer Science"
EXAM_MINUTES = 30              
TIME_WARNINGS = (10, 5, 1)     

PASSWORD_LENGTH = 4            
MAX_LOGIN_ATTEMPTS = 3

SPEECH_RATE = 1               
READ_INSTRUCTIONS_AT_START = True
ECHO_EACH_CHARACTER = True     
CONFIRM_SUBMIT = True          
ALLOW_INTERRUPT = True         
SEND_KEYS_TO_CBT = True        


CHUNK = 1000                   
STABLE_EXACT = 1               
STABLE_PREFIX = 2              

QUESTIONS = [
    {"text": "What is the capital of Nigeria?", "options": ["Lagos", "Abuja", "Ibadan", "Kaduna"]},
    {"text": "Which organ pumps blood through the body?", "options": ["Lungs", "Brain", "Heart", "Liver"]},
    {"text": "What is the primary functional component of an industrial computer processor?", "options": ["Transistor", "Capacitor", "Resistor", "Diode"]},
    {"text": "Which protocol handles secure encryption on the modern web?", "options": ["HTTP", "HTTPS", "FTP", "SMTP"]},
    {"text": "What is the standard structural framework layout language used for web applications?", "options": ["HTML", "Python", "C Plus Plus", "Java"]},
    {"text": "Which programming paradigm focuses strictly on modules and objects?", "options": ["Procedural", "Functional", "Object Oriented", "Logical"]},
    {"text": "What does RAM stand for in computer hardware architecture?", "options": ["Read Access Memory", "Random Access Memory", "Run Active Memory", "Rate Amplified Mod"]},
]

OPTION_LETTERS = "abcde"
NUMBER_WORDS = ["zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen", "twenty"]
WORD_TO_NUM = {w: i for i, w in enumerate(NUMBER_WORDS)}

LETTER_ALIASES = {
    "a": ["a", "ay", "eh"], "b": ["b", "bee", "be"], "c": ["c", "see", "sea"],
    "d": ["d", "dee"], "e": ["e"], "f": ["f", "eff"], "g": ["g", "gee"],
    "h": ["h", "aitch"], "i": ["i", "eye"], "j": ["j", "jay"], "k": ["k", "kay"],
    "l": ["l", "el"], "m": ["m", "em"], "n": ["n", "en"], "o": ["o", "oh"],
    "p": ["p", "pee"], "q": ["q", "queue", "cue"], "r": ["r", "are"],
    "s": ["s", "ess"], "t": ["t", "tea", "tee"], "u": ["u", "you"],
    "v": ["v", "vee"], "w": ["w", "double you", "double u"], "x": ["x", "ex"],
    "y": ["y", "why"], "z": ["z", "zed", "zee"],
}
DIGIT_WORDS = {
    "zero": "0", "one": "1", "two": "2", "three": "3", "four": "4",
    "five": "5", "six": "6", "seven": "7", "eight": "8", "nine": "9",
    "won": "1", "to": "2", "too": "2", "tree": "3", "for": "4", "fore": "4", "ate": "8",
}

SPELL_CHARS = dict(DIGIT_WORDS)
for _letter, _aliases in LETTER_ALIASES.items():
    for _alias in _aliases:
        if " " not in _alias:
            SPELL_CHARS[_alias] = _letter.upper()

COMMAND_PHRASES = {
    "next": ["next", "next question"],
    "previous": ["previous", "previous question", "back", "go back"],
    "repeat": ["repeat", "again", "read again", "repeat question"],
    "submit": ["submit", "submit exam", "finish"],
    "status": ["status", "review"],
    "time": ["time", "time left", "how much time"],
    "help": ["help", "instructions"],
    "faster": ["faster", "speak faster"],
    "slower": ["slower", "speak slower"],
    "louder": ["louder"],
    "quieter": ["quieter"],
}
PHRASE_TO_COMMAND = {p: c for c, phrases in COMMAND_PHRASES.items() for p in phrases}
GOTO_PHRASES = {}
for _n in range(1, len(QUESTIONS) + 1):
    for _form in (f"go to {NUMBER_WORDS[_n]}", f"go to question {NUMBER_WORDS[_n]}", f"question {NUMBER_WORDS[_n]}"):
        GOTO_PHRASES[_form] = _n

YES_WORDS = {"yes", "yeah", "okay", "ok", "fine", "good"}
CONFIRM_WORDS = {"confirm", "yes", "submit"}

SVSF_ASYNC = 1
SVSF_PURGE = 2
SVSF_XML = 8

voice = win32com.client.Dispatch("SAPI.SpVoice")
voices = voice.GetVoices()
voice_idx = 0
speech_rate = SPEECH_RATE
speech_volume = 100
voice.Rate = speech_rate
voice.Volume = speech_volume

def spell(char): return f"<spell>{char}</spell>"
def pause(ms=150): return f'<silence msec="{ms}"/>'
def strip_tags(text): return re.sub(r"<[^>]+>", "", text)

def speak(text, interruptible=False):
    """Speaks text. If interruptible is True, uses Async mode so we can kill it."""
    print(f"System: {strip_tags(text)}")
    flags = SVSF_XML
    if interruptible: flags |= SVSF_ASYNC
    try:
        voice.Speak(text, flags)
    except Exception as error:
        print(f"Speech error: {error}")

def stop_speaking():
    """Instantly kills current speech for zero-latency barge-in."""
    try:
        voice.Speak("", SVSF_ASYNC | SVSF_PURGE)
    except: pass

def change_rate(delta):
    global speech_rate
    speech_rate = max(-10, min(10, speech_rate + delta))
    voice.Rate = speech_rate
    speak("Faster." if delta > 0 else "Slower.")


if not os.path.exists(MODEL_PATH):
    print(f"Error: Vosk model not found in '{MODEL_PATH}'.")
    sys.exit(1)

SetLogLevel(VOSK_LOG_LEVEL)
print("Loading high-speed voice engine...")
model = Model(MODEL_PATH)

audio_queue = queue.Queue()

def audio_callback(indata, frames, time_info, status):
    if status: print(status, file=sys.stderr)
    audio_queue.put(bytes(indata))

stream = sd.RawInputStream(samplerate=16000, blocksize=CHUNK, dtype='int16',
                           channels=1, callback=audio_callback)
stream.start()

def flush_mic():
    """Empties the audio queue immediately to ignore past noise."""
    while not audio_queue.empty():
        try: audio_queue.get_nowait()
        except queue.Empty: break

def clean(text): return " ".join(w for w in text.split() if w != "[unk]")

def make_rec(words):
    rec = KaldiRecognizer(model, 16000, json.dumps(sorted(set(words)) + ["[unk]"]))
    try: rec.SetEndpointerMode(EndpointerMode.SHORT)
    except: pass
    return rec

class Listener:
    def __init__(self, phrases, partials=True):
        self.valid = set(phrases)
        self.partials = partials
        self.rec = make_rec(self.valid)
        self.prefixes = set(" ".join(phrase.split()[:k]) for phrase in self.valid for k in range(1, len(phrase.split())))
        self.last = ""
        self.stable = 0

    def reset(self):
        self.rec.Reset()
        self.last = ""
        self.stable = 0

    def feed(self, data):
        if self.rec.AcceptWaveform(data):
            text = clean(json.loads(self.rec.Result()).get("text", ""))
            self.last, self.stable = "", 0
            return text or None
        if not self.partials:
            return None
        
        partial = clean(json.loads(self.rec.PartialResult()).get("partial", ""))
        self.stable = self.stable + 1 if (partial and partial == self.last) else 0
        self.last = partial
        
        if partial in self.valid:
            needed = STABLE_PREFIX if partial in self.prefixes else STABLE_EXACT
            if self.stable >= needed:
                return partial
        return None

    def listen(self, clock=None, reading_question=False):
        """
        Listens continuously. 
        If reading_question is True, it will instantly stop SAPI upon hearing a valid command.
        """
        self.reset()
        flush_mic()
        
        while True:
            # Check clock
            if clock is not None:
                event = clock.tick()
                if event == "timeup": return "timeup"
                if event:
                    speak(event)
                    self.reset()
                    flush_mic()

            # Process audio chunks as fast as they arrive
            try:
                data = audio_queue.get(timeout=0.01) # Poll fast to allow clock ticks
                result = self.feed(data)
                if result:
                    if reading_question: stop_speaking() # BARGE-IN: instantly stop TTS
                    print(f"Heard: {result}")
                    return result
            except queue.Empty:
                # If we are reading async and it finishes naturally, just loop
                continue


listener_yes_no = Listener(YES_WORDS | {"no", "nope", "change"})
listener_start = Listener({"start"})
listener_confirm = Listener(CONFIRM_WORDS | {"cancel", "no", "back"})
spell_vocab = list(SPELL_CHARS) + ["double you", "double u", "delete", "clear"]
listener_spell = Listener(spell_vocab, partials=False)

def spell_tokens(text):
    return text.replace("double you", "w").replace("double u", "w").split()


class ExamClock:
    def __init__(self, minutes):
        self.minutes = minutes
        self.deadline = None
        self.pending = sorted([w for w in TIME_WARNINGS if minutes and w < minutes], reverse=True)

    def start(self):
        if self.minutes: self.deadline = time.time() + self.minutes * 60

    def remaining(self):
        return None if self.deadline is None else max(0, self.deadline - time.time())

    def tick(self):
        left = self.remaining()
        if left is None: return None
        if left <= 0: return "timeup"
        if self.pending and left <= self.pending[0] * 60:
            return f"{self.pending.pop(0)} minutes remaining."
        return None

    def spoken_left(self):
        left = self.remaining()
        if left is None: return "This exam has no time limit."
        if left < 60: return "Less than one minute remaining."
        return f"{int(left // 60)} minutes remaining."

def parse_reg_char(word, pos):
    char = SPELL_CHARS.get(word)
    if char is None: return None
    if pos == 0: return "U" if char == "U" else None
    if pos in (1, 2) or pos >= 5:
        if char == "O": char = "0"
        return char if char.isdigit() else None
    return char if char.isalpha() else None

def get_registration_number():
    chars = []
    speak("Spell your registration number.")
    while len(chars) < 9:
        for word in spell_tokens(listener_spell.listen()):
            if len(chars) >= 9: break
            if word == "clear": chars, _ = [], speak("Cleared.")
            elif word == "delete":
                if chars: chars.pop(), speak("Deleted.")
            else:
                char = parse_reg_char(word, len(chars))
                if char is None:
                    speak("Start with U." if not chars else "Letter expected." if len(chars) in (3,4) else "Number expected.")
                    break
                chars.append(char)
                if ECHO_EACH_CHARACTER: speak(spell(char), interruptible=True) # Async echo prevents stutter
    final_reg = "".join(chars)
    speak(f"{final_reg}. Confirmed.")
    return final_reg

def get_password():
    chars = []
    speak(f"Say your {PASSWORD_LENGTH} character password.")
    while len(chars) < PASSWORD_LENGTH:
        for word in spell_tokens(listener_spell.listen()):
            if len(chars) >= PASSWORD_LENGTH: break
            if word == "clear": chars, _ = [], winsound.Beep(400, 300)
            elif word == "delete":
                if chars: chars.pop(), winsound.Beep(500, 150)
            else:
                char = SPELL_CHARS.get(word)
                if char: chars.append(char), winsound.Beep(900, 150)
    return "".join(chars)


def cbt_select(letter):
    if SEND_KEYS_TO_CBT: pyautogui.press(letter)
def cbt_next():
    if SEND_KEYS_TO_CBT: pyautogui.press(["tab", "enter"])
def cbt_previous(): pass
def cbt_goto(number): pass

INSTRUCTIONS = (
    "Instructions. Answer by saying A, B, C or D, or the answer text. "
    "Say next for the next question, previous to go back. "
    "Say go to question 4 to jump. Say repeat to hear a question again. "
    "Say time or status to check progress. Say faster or slower to change speed. "
    "Say submit to finish."
)

def prepare_questions():
    print("Preparing questions...")
    for n, q in enumerate(QUESTIONS, start=1):
        q["num"] = n
        parts = [f"Question {n}.", escape(q["text"]), pause(200)]
        answers = {}
        option_text = {}
        for letter, option in zip(OPTION_LETTERS, q["options"]):
            option_text[letter] = option
            parts.append(f"{spell(letter.upper())}, {escape(option)}.{pause(100)}") # Faster pauses
            forms = [option.lower()]
            for alias in LETTER_ALIASES[letter]: forms += [alias, f"option {alias}"]
            for form in forms: answers[form] = letter
        q["spoken"] = " ".join(parts)
        q["answers"] = answers
        q["option_text"] = option_text
        q["listener"] = Listener(list(PHRASE_TO_COMMAND) + list(GOTO_PHRASES) + list(answers))

def question_text(q, responses):
    text = q["spoken"]
    if responses[q["num"]]:
        current = responses[q["num"]]
        text += f" Your answer is {spell(current)}, {escape(q['option_text'][current.lower()])}."
    return text

def speak_question(text, listener, clock):
    """Speaks the question and allows instant user interruption."""
    if not ALLOW_INTERRUPT:
        speak(text)
        return None
    
    speak(text, interruptible=True)
    
    action = listener.listen(clock, reading_question=True)
    
    stop_speaking()
    return action

def conduct_exam(reg_number, clock):
    total = len(QUESTIONS)
    responses = {q["num"]: None for q in QUESTIONS}
    idx = 0
    read_aloud = True
    action = None

    while True:
        q = QUESTIONS[idx]
        listener = q["listener"]

        if read_aloud:
            action = speak_question(question_text(q, responses), listener, clock)
            read_aloud = False
        else:
            action = listener.listen(clock)

        command = PHRASE_TO_COMMAND.get(action)

        if action == "timeup":
            speak("Time is up. Your exam has been submitted.")
            return responses
        elif command == "repeat": read_aloud = True
        elif command == "next":
            if idx == total - 1: speak("Last question. Say submit.")
            else: idx += 1; cbt_next(); read_aloud = True
        elif command == "previous":
            if idx == 0: speak("First question.")
            else: idx -= 1; cbt_previous(); read_aloud = True
        elif action in GOTO_PHRASES:
            idx = GOTO_PHRASES[action] - 1; cbt_goto(idx + 1); read_aloud = True
        elif command == "status":
            missing = [n for n, a in responses.items() if a is None]
            speak("All answered." if not missing else f"Unanswered: {', '.join(map(str, missing))}")
        elif command == "time": speak(clock.spoken_left())
        elif command == "help": speak(INSTRUCTIONS)
        elif command in ["faster", "slower"]: change_rate(2 if command=="faster" else -2)
        elif command == "submit":
            missing = [n for n, a in responses.items() if a is None]
            if missing: speak(f"You have {len(missing)} unanswered questions. Cancel submission to review.")
            else: speak("Say confirm to submit, or cancel.")
            
            if listener_confirm.listen(clock) in CONFIRM_WORDS:
                speak("Exam submitted. Thank you."); return responses
            else: speak("Cancelled.")
        elif action in q["answers"]:
            letter = q["answers"][action]
            cbt_select(letter)
            responses[q["num"]] = letter.upper()
            speak(f"{spell(letter.upper())}, {escape(q['option_text'][letter])}.")
        else:
            speak("Say that again.")

def run_exam_proctor():
    prepare_questions()
    
    speak("Welcome to the CBT system.")
    reg_number = get_registration_number()
    get_password()
    
    speak(f"You have {EXAM_MINUTES} minutes. Say start when ready.")
    while "start" not in listener_start.listen().split(): pass
    
    clock = ExamClock(EXAM_MINUTES)
    clock.start()
    speak("Exam started.")
    if READ_INSTRUCTIONS_AT_START: speak(INSTRUCTIONS)
    
    conduct_exam(reg_number, clock)

if __name__ == "__main__":
    try:
        run_exam_proctor()
    except KeyboardInterrupt:
        pass
    finally:
        stop_speaking()
        stream.stop()
        stream.close()