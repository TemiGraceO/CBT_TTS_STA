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


# ===============================================================
# CONFIGURATION
# ===============================================================
MODEL_PATH = "model"
VOSK_LOG_LEVEL = -1

COURSE_SPOKEN = "C S C two zero one"
COURSE_TITLE = "Introduction to Computer Science"
EXAM_MINUTES = 30
TIME_WARNINGS = (10, 5, 1)

PASSWORD_LENGTH = 4
MAX_LOGIN_ATTEMPTS = 3

SPEECH_RATE = 0                # -3 = very calm, 0 = normal, 2 = brisk. Adjust to taste.
READ_INSTRUCTIONS_AT_START = True
ECHO_EACH_CHARACTER = True
CONFIRM_SUBMIT = True
ALLOW_INTERRUPT = True         # Candidate can answer while a question is being read (use headphones)
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
}
PHRASE_TO_COMMAND = {p: c for c, phrases in COMMAND_PHRASES.items() for p in phrases}
GOTO_PHRASES = {}
for _n in range(1, len(QUESTIONS) + 1):
    for _form in (f"go to {NUMBER_WORDS[_n]}", f"go to question {NUMBER_WORDS[_n]}", f"question {NUMBER_WORDS[_n]}"):
        GOTO_PHRASES[_form] = _n

YES_WORDS = {"yes", "yeah", "okay", "ok", "fine", "good"}
CONFIRM_WORDS = {"confirm", "yes", "submit"}


# ===============================================================
# SPEECH OUTPUT (SAPI)
# ===============================================================
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
def pause(ms=200): return f'<silence msec="{ms}"/>'
def strip_tags(text): return re.sub(r"<[^>]+>", "", text)


def speak(text, interruptible=False):
    """Speaks text and waits until finished, unless interruptible is True."""
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
    except Exception:
        pass


def select_voice(index):
    global voice_idx
    voice_idx = index
    try:
        voice.Voice = voices.Item(index)
        voice.Rate = speech_rate
        voice.Volume = speech_volume
    except Exception as error:
        print(f"Voice error: {error}")


def change_rate(delta):
    global speech_rate
    new_rate = max(-10, min(10, speech_rate + delta))
    if new_rate == speech_rate:
        speak("This is already the fastest speed." if delta > 0 else "This is already the slowest speed.")
        return
    speech_rate = new_rate
    voice.Rate = speech_rate
    speak("Faster." if delta > 0 else "Slower.")


# ===============================================================
# ASYNC MICROPHONE (sounddevice)
# ===============================================================
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
    while not audio_queue.empty():
        try: audio_queue.get_nowait()
        except queue.Empty:
            break


def clean(text): return " ".join(w for w in text.split() if w != "[unk]")


def make_rec(words):
    rec = KaldiRecognizer(model, 16000, json.dumps(sorted(set(words)) + ["[unk]"]))
    try: rec.SetEndpointerMode(EndpointerMode.SHORT)
    except Exception: pass
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
        self.reset()
        flush_mic()
        while True:
            if clock is not None:
                event = clock.tick()
                if event == "timeup": return "timeup"
                if event:
                    speak(event)
                    self.reset()
                    flush_mic()

            try:
                data = audio_queue.get(timeout=0.01)
                result = self.feed(data)
                if result:
                    if reading_question: stop_speaking()  # BARGE-IN
                    print(f"Heard: {result}")
                    return result
            except queue.Empty:
                continue


# Listeners
listener_yes_no = Listener(YES_WORDS | {"no", "nope", "change"})
listener_start = Listener({"start", "begin", "start exam", "begin exam"})
listener_confirm = Listener(CONFIRM_WORDS | {"cancel", "no", "back"})
spell_vocab = list(SPELL_CHARS) + ["double you", "double u", "delete", "clear"]
listener_spell = Listener(spell_vocab, partials=False)


def spell_tokens(text):
    return text.replace("double you", "w").replace("double u", "w").split()


# ===============================================================
# VOICE SELECTION AT STARTUP
# ===============================================================
def choose_voice():
    if voices.Count < 2:
        speak("Only one voice is installed on this computer.")
        speak("Voice check. " + pause(200) + "This is how I will read your exam.")
        return

    speak("First, let us choose a voice. I will read a short sample.")
    for _ in range(voices.Count):
        speak(pause(70) + "Sample. This is how I will read your examination questions today.")
        speak("Is this voice okay? Say yes to keep it, or no for the next voice.")
        answer = listener_yes_no.listen()
        if answer in YES_WORDS:
            speak("Voice selected.")
            return
        select_voice((voice_idx + 1) % voices.Count)

    select_voice(0)
    speak("Returning to the first voice. Voice selected.")


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
    speak("Spell your registration number, one character at a time. Say delete to remove the last character, or clear to start over.")
    while len(chars) < 9:
        for word in spell_tokens(listener_spell.listen()):
            if len(chars) >= 9: break
            if word == "clear":
                chars = []
                speak("Cleared. Start again.")
            elif word == "delete":
                if chars: chars.pop(); speak("Deleted.")
            else:
                char = parse_reg_char(word, len(chars))
                if char is None:
                    speak("Start with U." if not chars else "Letter expected." if len(chars) in (3, 4) else "Number expected.")
                    break
                chars.append(char)
                if ECHO_EACH_CHARACTER:
                    speak(spell(char))  # synchronous: calm, paced, cannot self-trigger
    final_reg = "".join(chars)
    speak(f"{final_reg}. Confirmed.")
    return final_reg


def get_password():
    chars = []
    speak(f"Say your {PASSWORD_LENGTH} character password. You will hear a beep for each character.")
    while len(chars) < PASSWORD_LENGTH:
        for word in spell_tokens(listener_spell.listen()):
            if len(chars) >= PASSWORD_LENGTH: break
            if word == "clear": chars = []; winsound.Beep(400, 300)
            elif word == "delete":
                if chars: chars.pop(); winsound.Beep(500, 150)
            else:
                char = SPELL_CHARS.get(word)
                if char: chars.append(char); winsound.Beep(900, 150)
    return "".join(chars)


def verify_login(reg_number, password):
    # TODO: check against your student database.
    return True


def cbt_select(letter):
    if SEND_KEYS_TO_CBT: pyautogui.press(letter)


def cbt_next():
    if SEND_KEYS_TO_CBT: pyautogui.press(["tab", "enter"])


def cbt_previous(): pass
def cbt_goto(number): pass


# ===============================================================
# EXAM LOGIC
# ===============================================================
INSTRUCTIONS = (
    "Instructions. Answer by saying A, B, C or D, or the answer itself. "
    "Say next for the next question, and previous to go back. "
    "Say go to question 4, or any number, to jump to that question. "
    "Say repeat to hear a question again. "
    "Say status to hear which questions you have not answered. "
    "Say time to hear the time left. "
    "Say faster or slower to change my speed. "
    "Say help to hear these instructions again. "
    "Say submit when you are finished."
)


def prepare_questions():
    print("Preparing questions...")
    for n, q in enumerate(QUESTIONS, start=1):
        q["num"] = n
        parts = [f"Question {n}.", escape(q["text"]), pause(300)]
        answers = {}
        option_text = {}
        for letter, option in zip(OPTION_LETTERS, q["options"]):
            option_text[letter] = option
            parts.append(f"{spell(letter.upper())}, {escape(option)}.{pause(200)}")
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
    """Reads the question; candidate may interrupt by speaking a valid command."""
    if not ALLOW_INTERRUPT:
        speak(text)
        return None

    speak(text, interruptible=True)
    action = listener.listen(clock, reading_question=True)
    stop_speaking()  # stop the tail if it finished naturally
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

        elif command == "repeat":
            read_aloud = True

        elif command == "next":
            if idx == total - 1:
                speak("This is the last question. Say submit when you are done, or go to a question number.")
            else:
                idx += 1; cbt_next(); read_aloud = True

        elif command == "previous":
            if idx == 0:
                speak("This is the first question.")
            else:
                idx -= 1; cbt_previous(); read_aloud = True

        elif action in GOTO_PHRASES:
            idx = GOTO_PHRASES[action] - 1; cbt_goto(idx + 1); read_aloud = True

        elif command == "status":
            missing = [n for n, a in responses.items() if a is None]
            if not missing:
                speak(f"You have answered all {total} questions.")
            else:
                spoken = ", ".join(str(n) for n in missing)
                speak(f"You have {len(missing)} unanswered questions. Numbers {spoken}.")

        elif command == "time":
            speak(clock.spoken_left())

        elif command == "help":
            speak(INSTRUCTIONS)

        elif command == "faster":
            change_rate(2)
        elif command == "slower":
            change_rate(-2)

        elif command == "submit":
            missing = [n for n, a in responses.items() if a is None]
            if missing:
                # FIX: never offer submission while questions are unanswered.
                spoken = ", ".join(str(n) for n in missing)
                speak(f"You cannot submit yet. You have {len(missing)} unanswered questions. Numbers {spoken}. "
                      f"Say go to question, then the number.")
            else:
                speak("All questions answered. Say confirm to submit, or cancel to keep reviewing.")
                answer = listener_confirm.listen(clock)
                if answer == "timeup":
                    speak("Time is up. Your exam has been submitted.")
                    return responses
                if answer in CONFIRM_WORDS:
                    speak("Your exam has been submitted. Thank you.")
                    return responses
                speak("Submission cancelled.")

        elif action in q["answers"]:
            letter = q["answers"][action]
            cbt_select(letter)
            responses[q["num"]] = letter.upper()
            speak(f"{spell(letter.upper())}, {escape(q['option_text'][letter])}.")

        else:
            speak("Say that again.")


def run_exam_proctor():
    prepare_questions()

    speak("Welcome to the voice controlled examination system.")
    choose_voice()

    reg_number = None
    for _ in range(MAX_LOGIN_ATTEMPTS):
        reg_number = get_registration_number()
        password = get_password()
        if verify_login(reg_number, password):
            break
        speak("Login failed. Please try again.")
    else:
        speak("Too many failed attempts. Please call the invigilator.")
        return

    intro = f"You are about to write {COURSE_SPOKEN}, {COURSE_TITLE}. "
    if EXAM_MINUTES:
        intro += f"You have {EXAM_MINUTES} minutes. "
    speak(intro)

    if READ_INSTRUCTIONS_AT_START:
        speak(INSTRUCTIONS)

    # WAIT HERE: nothing happens until the candidate says start.
    speak("Take your time. Say start, when you are ready to begin.")
    while True:
        command = listener_start.listen()
        if command and any(word in command.split() for word in ("start", "begin")):
            break
        speak("Say start when you are ready.")

    clock = ExamClock(EXAM_MINUTES)
    clock.start()
    speak("The exam has started. Good luck.")
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