import os
import re
import sys
import json
import time
import string
from xml.sax.saxutils import escape

import winsound
import pyautogui
import pyaudio
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
VOSK_LOG_LEVEL = -1            # set to 0 to see Vosk messages when debugging

COURSE_SPOKEN = "C S C two zero one"          # how the course code is read aloud
COURSE_TITLE = "Introduction to Computer Science"
EXAM_MINUTES = 30              # set to None for no time limit
TIME_WARNINGS = (10, 5, 1)     # minutes remaining at which a warning is spoken

PASSWORD_LENGTH = 4            # letters and numbers allowed
MAX_LOGIN_ATTEMPTS = 3

SPEECH_RATE = 0                # -10 (slowest) to 10 (fastest)
READ_INSTRUCTIONS_AT_START = True
ECHO_EACH_CHARACTER = True     # speak each registration character as entered
CONFIRM_SUBMIT = True          # ask "confirm" before final submission
ALLOW_INTERRUPT = False        # True lets candidates answer while a question is being
                               # read. Use ONLY with headphones, or the microphone
                               # will hear the computer reading the options.
SEND_KEYS_TO_CBT = True        # press keys in the CBT window

CHUNK = 2000                   # audio frames per read (0.125 seconds)
STABLE_EXACT = 1               # chunks a phrase must stay unchanged before acting
STABLE_PREFIX = 4              # same, when the phrase could still grow ("next question")

# Write options the way they are spoken (words, not symbols).
QUESTIONS = [
    {"text": "What is the capital of Nigeria?",
     "options": ["Lagos", "Abuja", "Ibadan", "Kaduna"]},
    {"text": "Which organ pumps blood through the body?",
     "options": ["Lungs", "Brain", "Heart", "Liver"]},
    {"text": "What is the primary functional component of an industrial computer processor?",
     "options": ["Transistor", "Capacitor", "Resistor", "Diode"]},
    {"text": "Which protocol handles secure encryption on the modern web?",
     "options": ["HTTP", "HTTPS", "FTP", "SMTP"]},
    {"text": "What is the standard structural framework layout language used for web applications?",
     "options": ["HTML", "Python", "C Plus Plus", "Java"]},
    {"text": "Which programming paradigm focuses strictly on modules and objects?",
     "options": ["Procedural", "Functional", "Object Oriented", "Logical"]},
    {"text": "What does RAM stand for in computer hardware architecture?",
     "options": ["Read Access Memory", "Random Access Memory",
                 "Run Active Memory", "Rate Amplified Mod"]},
]

OPTION_LETTERS = "abcde"
NUMBER_WORDS = ["zero", "one", "two", "three", "four", "five", "six", "seven",
                "eight", "nine", "ten", "eleven", "twelve", "thirteen", "fourteen",
                "fifteen", "sixteen", "seventeen", "eighteen", "nineteen", "twenty"]
WORD_TO_NUM = {w: i for i, w in enumerate(NUMBER_WORDS)}
assert len(QUESTIONS) <= 20, "Navigation vocabulary supports up to 20 questions."

# Ways people (and accents) say each letter. Extra words improve recognition.
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

# word spoken -> character (uppercase letter or digit)
SPELL_CHARS = dict(DIGIT_WORDS)
for _letter, _aliases in LETTER_ALIASES.items():
    for _alias in _aliases:
        if " " not in _alias:
            SPELL_CHARS[_alias] = _letter.upper()

# Commands the candidate can say during the exam
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
    for _form in (f"go to {NUMBER_WORDS[_n]}",
                  f"go to question {NUMBER_WORDS[_n]}",
                  f"question {NUMBER_WORDS[_n]}"):
        GOTO_PHRASES[_form] = _n

YES_WORDS = {"yes", "yeah", "okay", "ok", "fine", "good"}
CONFIRM_WORDS = {"confirm", "yes", "submit"}


# ===============================================================
# SPEECH OUTPUT (Windows speech engine, runs inside Python)
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


def spell(char):
    """Makes the voice say a letter by its name."""
    return f"<spell>{char}</spell>"


def pause(ms=200):
    return f'<silence msec="{ms}"/>'


def strip_tags(text):
    return re.sub(r"<[^>]+>", "", text)


def speak(text):
    """Speaks text and waits until finished. Text may use SAPI tags."""
    print(f"System: {strip_tags(text)}")
    try:
        voice.Speak(text, SVSF_XML)
    except Exception as error:
        print(f"Speech error: {error}")


def select_voice(index):
    voice.Voice = voices.Item(index)


def change_rate(delta):
    global speech_rate
    new_rate = max(-10, min(10, speech_rate + delta))
    if new_rate == speech_rate:
        speak("This is already the fastest speed." if delta > 0
              else "This is already the slowest speed.")
        return
    speech_rate = new_rate
    voice.Rate = speech_rate
    speak("Faster." if delta > 0 else "Slower.")


def change_volume(delta):
    global speech_volume
    new_volume = max(20, min(100, speech_volume + delta))
    if new_volume == speech_volume:
        speak("This is already the loudest." if delta > 0 else "This is already the quietest.")
        return
    speech_volume = new_volume
    voice.Volume = speech_volume
    speak("Louder." if delta > 0 else "Quieter.")


# ===============================================================
# SPEECH INPUT (Vosk)
# ===============================================================
if not os.path.exists(MODEL_PATH):
    print(f"Error: Please download the Vosk model and place it in the '{MODEL_PATH}' folder.")
    sys.exit(1)

SetLogLevel(VOSK_LOG_LEVEL)
print("Loading voice engine, please wait...")
model = Model(MODEL_PATH)

mic = pyaudio.PyAudio()
stream = mic.open(format=pyaudio.paInt16, channels=1, rate=16000,
                  input=True, frames_per_buffer=CHUNK)
stream.start_stream()


def flush_mic():
    available = stream.get_read_available()
    if available:
        stream.read(available, exception_on_overflow=False)


def clean(text):
    return " ".join(w for w in text.split() if w != "[unk]")


def make_rec(words):
    """Recognizer limited to the given phrases. [unk] lets noise be ignored."""
    rec = KaldiRecognizer(model, 16000, json.dumps(sorted(set(words)) + ["[unk]"]))
    try:
        rec.SetEndpointerMode(EndpointerMode.SHORT)   # respond sooner after speech ends
    except Exception:
        pass
    return rec


class Listener:
    """Listens for a fixed set of phrases.

    With partials=True it acts as soon as a complete valid phrase is recognized,
    without waiting for silence. With partials=False (used for spelling) it
    waits for the end of the utterance so no spoken characters are lost.
    """

    def __init__(self, phrases, partials=True):
        self.valid = set(phrases)
        self.partials = partials
        self.rec = make_rec(self.valid)
        self.prefixes = set()
        for phrase in self.valid:
            parts = phrase.split()
            for k in range(1, len(parts)):
                self.prefixes.add(" ".join(parts[:k]))
        self.last = ""
        self.stable = 0

    def reset(self):
        self.rec.Reset()
        self.last = ""
        self.stable = 0

    def feed(self, data):
        rec = self.rec
        if rec.AcceptWaveform(data):
            text = clean(json.loads(rec.Result()).get("text", ""))
            self.last, self.stable = "", 0
            return text or None
        if not self.partials:
            return None
        partial = clean(json.loads(rec.PartialResult()).get("partial", ""))
        self.stable = self.stable + 1 if (partial and partial == self.last) else 0
        self.last = partial
        if partial in self.valid:
            needed = STABLE_PREFIX if partial in self.prefixes else STABLE_EXACT
            if self.stable >= needed:
                return partial
        return None

    def listen(self, clock=None):
        time.sleep(0.1)          # let the speaker echo fade
        self.reset()
        flush_mic()
        while True:
            if clock is not None:
                event = clock.tick()
                if event == "timeup":
                    return "timeup"
                if event:
                    speak(event)
                    time.sleep(0.1)
                    self.reset()
                    flush_mic()
            result = self.feed(stream.read(CHUNK, exception_on_overflow=False))
            if result:
                print(f"Heard: {result}")
                return result


# Listeners used outside the questions
listener_yes_no = Listener(YES_WORDS | {"no", "nope", "change"})
listener_start = Listener({"start"})
listener_confirm = Listener(CONFIRM_WORDS | {"cancel", "no", "back"})
spell_vocab = list(SPELL_CHARS) + ["double you", "double u", "delete", "clear"]
listener_spell = Listener(spell_vocab, partials=False)


def spell_tokens(text):
    text = text.replace("double you", "w").replace("double u", "w")
    return text.split()


# ===============================================================
# EXAM CLOCK
# ===============================================================
class ExamClock:
    def __init__(self, minutes):
        self.minutes = minutes
        self.deadline = None
        self.pending = sorted([w for w in TIME_WARNINGS if minutes and w < minutes],
                              reverse=True)

    def start(self):
        if self.minutes:
            self.deadline = time.time() + self.minutes * 60

    def remaining(self):
        return None if self.deadline is None else max(0, self.deadline - time.time())

    def tick(self):
        """Returns 'timeup', a warning to speak, or None."""
        left = self.remaining()
        if left is None:
            return None
        if left <= 0:
            return "timeup"
        if self.pending and left <= self.pending[0] * 60:
            minutes = self.pending.pop(0)
            return f"{minutes} {'minute' if minutes == 1 else 'minutes'} remaining."
        return None

    def spoken_left(self):
        left = self.remaining()
        if left is None:
            return "This exam has no time limit."
        if left < 60:
            return "Less than one minute remaining."
        minutes = int(left // 60)
        return f"{minutes} {'minute' if minutes == 1 else 'minutes'} remaining."


# ===============================================================
# SETUP AND LOGIN
# ===============================================================
def choose_voice():
    global voice_idx
    speak("Welcome to the examination system. Is this voice okay? Say yes or no.")
    for _ in range(max(1, voices.Count)):
        if listener_yes_no.listen() in YES_WORDS:
            return
        if voices.Count < 2:
            speak("No other voices are installed on this computer.")
            return
        voice_idx = (voice_idx + 1) % voices.Count
        select_voice(voice_idx)
        speak("Is this voice okay? Say yes or no.")


def parse_reg_char(word, pos):
    """Registration pattern: U, 2 numbers, 2 letters, 4 numbers (9 characters)."""
    char = SPELL_CHARS.get(word)
    if char is None:
        return None
    if pos == 0:
        return "U" if char == "U" else None
    if pos in (1, 2) or pos >= 5:
        if char == "O":
            char = "0"
        return char if char.isdigit() else None
    return char if char.isalpha() else None


def get_registration_number():
    chars = []
    speak("Spell your registration number, one character at a time.")
    while len(chars) < 9:
        for word in spell_tokens(listener_spell.listen()):
            if len(chars) >= 9:
                break
            if word == "clear":
                chars = []
                speak("Cleared. Start again.")
                continue
            if word == "delete":
                if chars:
                    chars.pop()
                    speak("Deleted.")
                continue
            char = parse_reg_char(word, len(chars))
            if char is None:
                if not chars:
                    speak("Start with U.")
                elif len(chars) in (3, 4):
                    speak("Expected a letter.")
                else:
                    speak("Expected a number.")
                break   # drop the rest of this utterance so positions stay aligned
            chars.append(char)
            if ECHO_EACH_CHARACTER:
                speak(spell(char))
    speak(pause(150).join(spell(c) for c in chars) + ". Confirmed.")
    return "".join(chars)


def get_password():
    """Letters and numbers. A beep per character, nothing spoken aloud."""
    chars = []
    speak(f"Say your {PASSWORD_LENGTH} character password. Letters and numbers.")
    while len(chars) < PASSWORD_LENGTH:
        for word in spell_tokens(listener_spell.listen()):
            if len(chars) >= PASSWORD_LENGTH:
                break
            if word == "clear":
                chars = []
                winsound.Beep(400, 300)
                continue
            if word == "delete":
                if chars:
                    chars.pop()
                    winsound.Beep(500, 150)
                continue
            char = SPELL_CHARS.get(word)
            if char is None:
                break
            chars.append(char)
            winsound.Beep(900, 150)
    return "".join(chars)


def verify_login(reg_number, password):
    # TODO: check reg_number and password against your student database.
    return True


# ===============================================================
# CBT SCREEN HOOKS (adapt these when you attach to the CBT software)
# ===============================================================
def cbt_select(letter):
    if SEND_KEYS_TO_CBT:
        pyautogui.press(letter)


def cbt_next():
    if SEND_KEYS_TO_CBT:
        pyautogui.press("tab")
        pyautogui.press("enter")


def cbt_previous():
    pass   # TODO: key sequence for "previous question" in your CBT software


def cbt_goto(number):
    pass   # TODO: key sequence for jumping to a question in your CBT software


# ===============================================================
# EXAM
# ===============================================================
INSTRUCTIONS = (
    f"Instructions. To answer, say {spell('A')}, {spell('B')}, {spell('C')} or {spell('D')}, "
    "or say the answer itself. "
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
            parts.append(f"{spell(letter.upper())}, {escape(option)}.{pause(250)}")
            forms = [option.lower()]
            if option.isupper() and len(option) > 1:
                forms.append(" ".join(option.lower()))        # "h t t p"
            for alias in LETTER_ALIASES[letter]:
                forms += [alias, f"option {alias}"]
            for form in forms:
                answers[form] = letter
        q["spoken"] = " ".join(parts)
        q["answers"] = answers
        q["option_text"] = option_text
        vocab = list(PHRASE_TO_COMMAND) + list(GOTO_PHRASES) + list(answers)
        q["listener"] = Listener(vocab)


def question_text(q, responses):
    text = q["spoken"]
    current = responses[q["num"]]
    if current:
        text += (f" Your answer is {spell(current)}, "
                 f"{escape(q['option_text'][current.lower()])}.")
    return text


def speak_question(text, listener):
    """Reads a question. With ALLOW_INTERRUPT, returns a command heard while reading."""
    if not ALLOW_INTERRUPT:
        speak(text)
        return None
    print(f"System: {strip_tags(text)}")
    flush_mic()
    listener.reset()
    heard = None
    try:
        voice.Speak(text, SVSF_ASYNC | SVSF_XML)
        while not voice.WaitUntilDone(0):
            phrase = listener.feed(stream.read(CHUNK, exception_on_overflow=False))
            if phrase:
                heard = phrase
                break
        if heard:
            voice.Speak("", SVSF_ASYNC | SVSF_PURGE)
    except Exception as error:
        print(f"Speech error: {error}")
    return heard


def join_numbers(numbers):
    items = [str(n) for n in numbers]
    if len(items) == 1:
        return items[0]
    if len(items) == 2:
        return f"{items[0]} and {items[1]}"
    return ", ".join(items[:-1]) + f", and {items[-1]}"


def unanswered(responses):
    return [n for n, answer in responses.items() if answer is None]


def status_text(responses):
    missing = unanswered(responses)
    total = len(responses)
    if not missing:
        return f"You have answered all {total} questions."
    return (f"You have answered {total - len(missing)} of {total}. "
            f"Unanswered: {join_numbers(missing)}.")


def handle_submit(responses, clock):
    """Returns True when the exam is finally submitted."""
    missing = unanswered(responses)
    if missing:
        count = len(missing)
        noun = "question" if count == 1 else "questions"
        label = "Number" if count == 1 else "Numbers"
        speak(f"You have {count} {noun} unanswered. {label} {join_numbers(missing)}. "
              f"Say go to question, then the number.")
        return False
    if CONFIRM_SUBMIT:
        speak("All questions answered. Say confirm to submit, or cancel to keep reviewing.")
        answer = listener_confirm.listen(clock)
        if answer == "timeup":
            speak("Time is up. Your exam has been submitted.")
            return True
        if answer not in CONFIRM_WORDS:
            speak("Submission cancelled.")
            return False
    speak("Your exam has been submitted. Thank you.")
    return True


def save_progress(reg_number, responses, submitted=False):
    try:
        with open(f"result_{reg_number}.json", "w") as f:
            json.dump({
                "registration_number": reg_number,
                "course": f"{COURSE_SPOKEN} {COURSE_TITLE}",
                "answers": responses,
                "submitted": submitted,
                "saved_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            }, f, indent=2)
    except OSError as error:
        print(f"Could not save results: {error}")


def conduct_exam(reg_number, clock):
    total = len(QUESTIONS)
    responses = {q["num"]: None for q in QUESTIONS}
    idx = 0
    read_aloud = True
    pending = None

    while True:
        q = QUESTIONS[idx]
        listener = q["listener"]

        if read_aloud:
            pending = speak_question(question_text(q, responses), listener)
            read_aloud = False

        action = pending or listener.listen(clock)
        pending = None
        command = PHRASE_TO_COMMAND.get(action)

        if action == "timeup":
            speak("Time is up. Your exam has been submitted.")
            save_progress(reg_number, responses, submitted=True)
            return responses

        elif command == "repeat":
            read_aloud = True

        elif command == "next":
            if idx == total - 1:
                speak("This is the last question. Say submit when you are done, "
                      "or go to a question number.")
            else:
                idx += 1
                cbt_next()
                read_aloud = True

        elif command == "previous":
            if idx == 0:
                speak("This is the first question.")
            else:
                idx -= 1
                cbt_previous()
                read_aloud = True

        elif action in GOTO_PHRASES:
            idx = GOTO_PHRASES[action] - 1
            cbt_goto(idx + 1)
            read_aloud = True

        elif command == "status":
            speak(status_text(responses))

        elif command == "time":
            speak(clock.spoken_left())

        elif command == "help":
            speak(INSTRUCTIONS)

        elif command == "faster":
            change_rate(2)
        elif command == "slower":
            change_rate(-2)
        elif command == "louder":
            change_volume(20)
        elif command == "quieter":
            change_volume(-20)

        elif command == "submit":
            if handle_submit(responses, clock):
                save_progress(reg_number, responses, submitted=True)
                return responses

        elif action in q["answers"]:
            letter = q["answers"][action]
            cbt_select(letter)
            responses[q["num"]] = letter.upper()
            save_progress(reg_number, responses)
            speak(f"{spell(letter.upper())}, {escape(q['option_text'][letter])}.")

        else:
            speak("Say that again.")


def run_exam_proctor():
    prepare_questions()
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
    if READ_INSTRUCTIONS_AT_START:
        intro += INSTRUCTIONS + " "
    speak(intro + "Say start when you are ready.")

    while "start" not in listener_start.listen().split():
        pass

    clock = ExamClock(EXAM_MINUTES)
    clock.start()
    speak("Exam started.")
    conduct_exam(reg_number, clock)


if __name__ == "__main__":
    try:
        run_exam_proctor()
    except KeyboardInterrupt:
        print("\nSystem shut down safely.")
    finally:
        stream.stop_stream()
        stream.close()
        mic.terminate()