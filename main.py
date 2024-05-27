import win32com.client as wincom
from googletrans import Translator
import speech_recognition as sr
import re
import sys
import pyttsx3
# For windows
speak = wincom.Dispatch("SAPI.SpVoice")

lang_list = {
   
'Afrikaans' : 'af',
'Albanian' : 'sq',
'Amharic' : 'am',
'Arabic' : 'ar',
'Armenian' : 'hy',
'Azerbaijani' : 'az',
'Basque' : 'eu',
'Belarusian' : 'be',
'Bengali' : 'bn',
'Bosnian' : 'bs',
'Bulgarian' : 'bg',
'Catalan' : 'ca',
'Cebuano':'ceb',
'Chichewa' : 'ny',
'Chinese':'zh-CN',
'Corsican' : 'co',
'Croatian' : 'hr',
'Czech' : 'cs',
'Danish' : 'da',
'Dutch' : 'nl',
'English' : 'en',
'Esperanto' : 'eo',
'Estonian' : 'et',
'Filipino' : 'tl',
'Finnish' : 'fi',
'French' : 'fr',
'Frisian' : 'fy',
'Galician' : 'gl',
'Georgian' : 'ka',
'German' : 'de',
'Greek' : 'el',
'Gujarati' : 'gu',
'Haitian Creole' : 'ht',
'Hausa' : 'ha',
'Hawaiian ':'haw',
'Hebrew' : 'he',
'Hindi' : 'hi',
'Hmong ' : 'hmn',
'Hungarian' : 'hu',
'Icelandic' : 'is',
'Igbo' : 'ig',
'Indonesian' : 'id',
'Irish' : 'ga',
'Italian' : 'it',
'Japanese' : 'ja',
'Javanese' : 'jw',
'Kannada' : 'kn',
'Kazakh' : 'kk',
'Khmer' : 'km',
'Kinyarwanda' : 'rw',
'Korean' : 'ko',
'Kurdish' : 'ku',
'Kyrgyz' : 'ky',
'Lao' : 'lo',
'Latin' : 'la',
'Latvian' : 'lv',
'Lithuanian' : 'lt',
'Luxembourgish' : 'lb',
'Macedonian' : 'mk',
'Malagasy' : 'mg',
'Malay' : 'ms',
'Malayalam' : 'ml',
'Maltese' : 'mt',
'Maori' : 'mi',
'Marathi' : 'mr',
'Mongolian' : 'mn',
'Myanmar ' : 'my',
'Nepali' : 'ne',
'Norwegian' : 'no',
'Odia' : 'or',
'Pashto' : 'ps',
'Persian' : 'fa',
'Polish' : 'pl',
'Portuguese' : 'pt',
'Punjabi' : 'pa',
'Romanian' : 'ro',
'Russian' : 'ru',
'Samoan' : 'sm',
'Scots Gaelic' : 'gd',
'Serbian' : 'sr',
'Sesotho' : 'st',
'Shona' : 'sn',
'Sindhi' : 'sd',
'Sinhala' : 'si',
'Slovak' : 'sk',
'Slovenian': 'sl',
'Somali' : 'so',
'Spanish': 'es',
'Sundanese': 'su',
'Swahili ': 'sw',
'Swedish' : 'sv',
'Tajik' : 'tg',
'Tamil' : 'ta',
'Tatar ': 'tt',
'Telugu' : 'te',
'Thai ': 'th',
'Turkish ': 'tr',
'Turkmen' : 'tk',
'Ukrainian ': 'uk',
'Urdu' : 'ur',
'Uyghur' : 'ug',
'Uzbek' : 'uz',
'Vietnamese' : 'vi',
'Welsh' : 'cy',
'Xhosa' : 'xh',
'Yiddish' : 'yi',
'Yoruba' : 'yo',
'Zulu' : 'zu',

}
def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        # r.pause_threshold =  0.6
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query.lower()
        except Exception as e:
            return "ErrorOcurred"
def translate(query):

    # Initialize the translator
    translator = Translator()

    # Detect language
    text = query
    detected = translator.detect(text)
    print(f"Detected language: {detected.lang}")
    out_lang = input("Which Language you'd like to translate: ")
    lang = lang_list[out_lang.title()]
    # Translate text
    translated = translator.translate(text, dest=lang)
    print(f"Translated text: {translated.text}")

    # Initialize the text-to-speech engine
    engine = pyttsx3.init()

    # Set properties (optional)
    engine.setProperty('rate', 150)  # Speed of speech (words per minute)
    engine.setProperty('volume', 1.0)  # Volume (0.0 to 1.0)

    # Speak text in English

    # Change the language and speak in French
    engine.setProperty('language', lang)  # Set language to French
    engine.say(translated.text)
    engine.runAndWait()


def main():
    speak.Speak("Hey, Tell me something which you'd like to translate")
    while True:  
        query = takeCommand() 
        translate(query)
        try:
            quit = re.search(r'.+ quit',query)
            if quit:
                speak.Speak('See you soon')
                raise TimeoutError("Tranza Quit")
        except:
            sys.exit()
    
if __name__ == "__main__":
    main()