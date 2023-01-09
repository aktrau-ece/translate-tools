'''
Subtitle Translator.py [2023.1.07]

Requirements:
	Python

Limitations
	Wont work for hardcoded subs
	Sentences which span multiple time-ranges won't be translated accurately (subtitles for each time-range are translated individually, which can sometimes mean that theyre translated out of context.)
'''

# CONFIG ==========================================

from types import SimpleNamespace
cfg = SimpleNamespace(**{
	# ** Paths can be relative and must end with "\\" **
	"inPath": "InputFolder\\",
	"interPath": "IntermediateFolder\\",
	"outPath": "OutputFolder\\",

	"inLang": "en", # input language
	"outLang": ["zh"], # list containing output languages
		# translation to multiple languages isnt done yet but i havent found a need for it so maybe just change the for loop

	"translateModels": {
		# dictionary, whose:
			# keys represent: folders containing pre-trained opus-mt translation models (folder names must end with "\\")
			# values represent: a list containing the source and target language, respectively
		"opus-mt-en-fr\\": ["en", "fr"],
		"opus-mt-en-de\\": ["en", "de"],
		"opus-mt-en-it\\": ["en", "it"],
		"opus-mt-en-es\\": ["en", "es"],
		"opus-mt-en-ar\\": ["en", "ar"],
		"opus-mt-en-vi\\": ["en", "vi"],
		"opus-mt-en-zh\\": ["en", "zh"],
	},

	"verbosity": 4,
	"testingMode": True,
})
print("Config options:", str(vars(cfg)))

# MODULES ========================================

try:
	print("Importing Modules..")
	import os # to send operating system commands, and work with os files
	import io
	import shutil # copying files
	import sys

	from easynmt import EasyNMT, models # NMT translator

except Exception as e: 
	print("Error when importing modules: " + str(e))
	raise SystemExit

# DATA ============================================

 # from: https://py-googletrans.readthedocs.io/en/latest/
LANGUAGES = { # this is really just a dictionary showing languages and their abbreviations. you need pre-trained models in order to translate
    'af': 'afrikaans',
    'sq': 'albanian',
    'am': 'amharic',
    'ar': 'arabic',
    'hy': 'armenian',
    'az': 'azerbaijani',
    'eu': 'basque',
    'be': 'belarusian',
    'bn': 'bengali',
    'bs': 'bosnian',
    'bg': 'bulgarian',
    'ca': 'catalan',
    'ceb': 'cebuano',
    'ny': 'chichewa',
    'zh-cn': 'chinese (simplified)',
    'zh-tw': 'chinese (traditional)',
    'co': 'corsican',
    'hr': 'croatian',
    'cs': 'czech',
    'da': 'danish',
    'nl': 'dutch',
    'en': 'english',
    'eo': 'esperanto',
    'et': 'estonian',
    'tl': 'filipino',
    'fi': 'finnish',
    'fr': 'french',
    'fy': 'frisian',
    'gl': 'galician',
    'ka': 'georgian',
    'de': 'german',
    'el': 'greek',
    'gu': 'gujarati',
    'ht': 'haitian creole',
    'ha': 'hausa',
    'haw': 'hawaiian',
    'iw': 'hebrew',
    'he': 'hebrew',
    'hi': 'hindi',
    'hmn': 'hmong',
    'hu': 'hungarian',
    'is': 'icelandic',
    'ig': 'igbo',
    'id': 'indonesian',
    'ga': 'irish',
    'it': 'italian',
    'ja': 'japanese',
    'jw': 'javanese',
    'kn': 'kannada',
    'kk': 'kazakh',
    'km': 'khmer',
    'ko': 'korean',
    'ku': 'kurdish (kurmanji)',
    'ky': 'kyrgyz',
    'lo': 'lao',
    'la': 'latin',
    'lv': 'latvian',
    'lt': 'lithuanian',
    'lb': 'luxembourgish',
    'mk': 'macedonian',
    'mg': 'malagasy',
    'ms': 'malay',
    'ml': 'malayalam',
    'mt': 'maltese',
    'mi': 'maori',
    'mr': 'marathi',
    'mn': 'mongolian',
    'my': 'myanmar (burmese)',
    'ne': 'nepali',
    'no': 'norwegian',
    'or': 'odia',
    'ps': 'pashto',
    'fa': 'persian',
    'pl': 'polish',
    'pt': 'portuguese',
    'pa': 'punjabi',
    'ro': 'romanian',
    'ru': 'russian',
    'sm': 'samoan',
    'gd': 'scots gaelic',
    'sr': 'serbian',
    'st': 'sesotho',
    'sn': 'shona',
    'sd': 'sindhi',
    'si': 'sinhala',
    'sk': 'slovak',
    'sl': 'slovenian',
    'so': 'somali',
    'es': 'spanish',
    'su': 'sundanese',
    'sw': 'swahili',
    'sv': 'swedish',
    'tg': 'tajik',
    'ta': 'tamil',
    'te': 'telugu',
    'th': 'thai',
    'tr': 'turkish',
    'uk': 'ukrainian',
    'ur': 'urdu',
    'ug': 'uyghur',
    'uz': 'uzbek',
    'vi': 'vietnamese',
    'cy': 'welsh',
    'xh': 'xhosa',
    'yi': 'yiddish',
    'yo': 'yoruba',
    'zu': 'zulu'}

# FUNCTIONS =======================================

class G:
	# Given a directory path (string), returns a list of filenames (strings) (with extensions)
	def listFiles(dir):
		full_list = os.listdir(dir)
		return full_list

	# Returns a files extension, including the dot
	def extension(fileName):
		for ch in range(len(fileName)-1, -1, -1):
			if fileName[ch] == "." :
				return fileName[ch:]
		return ""

	# Returns a file title (filename without the extension)
	def basename(filename):
		for ch in range(len(filename)-1, -1, -1):
			if filename[ch] == "." :
				return filename[:ch]
		return filename

	def wrap(root, affix):
		return affix + root + affix

	# Prints multidimensional arrays nicely by indenting based on layer
	def printArray(array, i = 0): # optional parameter 'i' specifies starting indentation layer. it is incremented as the function recurses
		for s in array:
			if isinstance(s, list): # https://stackoverflow.com/questions/26544091/checking-if-type-list-in-python
				printArray(s, i+1)
			else: print("\t"*i + s) # print as many tabs as is specified by 'i', then the object

	# Similar to printArray, just for dictionaries
	def printDict(dictionary, indentLevel = 0):
		for u in dictionary.keys():
			print('\t'*indentLevel + u)
			if type(dictionary[u]) == dict:
				printDict(dictionary[u], indentLevel+1)
			elif type(dictionary[u]) == str:
				print('\t'*(indentLevel+1) + dictionary[u])

	def showErr(userMsg = "Error", reason = ""):
		if reason == "": print(userMsg)
		else: print(userMsg + ":", reason)
		input("Press <ENTER> to exit")
		raise SystemExit

class Translate:
	class ModelOrganiser: # data type containing translation models, and their input and output languages
		def __init__(self):
			self.repo = {} # 2D dictionary whose first dimension represents source languages, second dimension represents target languages

		def add(self, path, sourceLang, targetLang):
			if not (sourceLang in self.repo): self.repo[sourceLang] = {}
			if targetLang in self.repo[sourceLang]: G.showErr(reason="Attempted to load a translation model whose source and target language already exists in ModelOrganiser()")
			else: self.repo[sourceLang][targetLang] = EasyNMT(translator=models.AutoModel(path))

		def get(self, sourceLang, targetLang):
			if not(sourceLang in self.repo) or not(targetLang in self.repo[sourceLang]): G.showErr(reason="Attemped to use a translation model whose translation direction doesnt exist in ModelOrganiser()")
			else: return self.repo[sourceLang][targetLang]

	translateModels = ModelOrganiser() # class variable containing a ModelOrganiser()

	def load_translateModels(): # for all the selected output languages, load the appropriate translation model from its folder
		print("Loading translation models...")
		for key in cfg.translateModels:
			if (cfg.translateModels[key][0] == cfg.inLang) and (cfg.translateModels[key][1] in cfg.outLang):
				Translate.translateModels.add(key, cfg.translateModels[key][0], cfg.translateModels[key][1])

	def translateText(text, sourceLang, targetLang):
		return Translate.translateModels.get(sourceLang, targetLang).translate(text, source_lang=sourceLang, target_lang=targetLang, max_new_tokens=512)

	def translateText_robust(text, sourceLang, targetLang):
		if cfg.verbosity >= 4: print("Translating text: " + G.wrap(text, "\""))

		#if "\n" in text: return Translate.translateText_robust(text.replace("\n", " "), sourceLang, targetLang)

		# some characters / combinations of characters will make the translater return some wacky stuff.
		# known cases:
			# strings containing only numbers (and perhaps spaces and periods)
			# the '≈' symbol on its own
			# "NUMBER. (SOME_TEXT)" this will cause the translator to go on about the european council and some other bullshit
			# greek letters, when translated alone, make the translator output nonsense
			# the letter 'r' (by itself) . It makes the translator spit out some garbage. Im going to pre-emptively avoid translating any single letters.
			# "TEXT!)." the translator returns the text, exclamation mark, and closed parenthesis followed by a bunch of periods
		# {
		if (text is None) or (all([(v == " " or v == "\t") for v in text])): res = ""
		elif all([(v.isnumeric() | (v==" ") | (v=="\t") | (v==".")) for v in text]): res = text # if every char in the string is either a number, a tab, a space, or a period
		elif text == "≈": res = text # if the string contains only '≈'

		elif (len(text) > 2) and (text[0].isnumeric()) & (text[1] == "."):
			if cfg.verbosity >= 5: print("Recursing")
			res = text[0:2] + Translate.translateText_robust(text[2:], sourceLang, targetLang)
				
		elif (len(text) == 1): res = text
		
		elif (text.find("!).") != -1) and (all([(v == ".") for v in text[text.find("!).")+2:]])):
			res = (Translate.translateText_robust(text[:text.find("!).")+2], sourceLang, targetLang) + str(text[text.find("!).")+2:]))
		# }

		else:
			res = Translate.translateText(text, sourceLang, targetLang)
			if cfg.verbosity >= 5: print("Clear")

		if (
			len(res) > 3*len(text)
			or (".........." in res)
			or ("----------" in res)
		):
			if cfg.verbosity >= 4: print("Translation looks like some garbage. Using original un-translated text..")
			return text # if the translator bugs and returns a bunch of garbage, return the original untranslated text
		else: return res

# DATA TYPES =======================================

class Subtitle:
	def __init__(self, number, timeRange, text=[]):
		self.number = number
		self.timeRange = timeRange
		self.text = text

	def __str__(self): return (self.number + "\n" + self.timeRange + "\n" + self.text + "\n")

# MAIN =============================================

# discover files
fileList = SimpleNamespace(**{})
fileList.input = G.listFiles(cfg.inPath)
fileList.input_srt = [f for f in fileList.input if G.extension(f) == ".srt"]
print("Found " + str(len(fileList.input_srt)) + " .srt file(s).")

# initialize translator
Translate.load_translateModels()

def _translateSubtitles(filename):
	subs_original = [] # list of untranslated subtitles
	subs_translated = [] # list of translated subtitles

	# open and read the file {
	try:
		print("Opening " + filename + " with utf-8 encoding...", end="")
		file_read = io.open(cfg.inPath + filename, mode="r", encoding="utf-8")
		lines = file_read.readlines() # array of strings, each representing a line, and ending with "\n"
		print(" success")

	except Exception as e:
		print(" fail")

		try:
			print("Opening " + filename + " with utf-16 encoding...", end="")
			file_read = io.open(cfg.inPath + filename, mode="r", encoding="utf-16")
			lines = file_read.readlines()
			print(" success")

		except Exception as e_1:
			print(" fail")
			print("Exiting..."); raise SystemExit

	#if cfg.testingMode: print("\n" + str(lines))
	# }

	# Expected syntax for .srt files:
	# 2
	# 00:00:34,333 --> 00:00:39,750
	# A long time ago, it is said,
	# a monster came here.

	print()

	# load all subs {
	hold = [None, None, None] # temporary hold for a subtitle. [0] is for a subititles' index number, [1] is for its time range, and [2] is for its text.

	for l, line in enumerate(lines):
		line = line.replace("\n", "") # remove the new-line character which is at the end of every line
		if cfg.verbosity >= 5: print(f'Line {l}: ' + G.wrap(line, "\""))

		# check if the line is empty. this can mark the end of a subtitle, or can be at the beginning or end of a file.
		if line == "":
			if all([v == None for v in hold]): # if all values in hold are == None
				print("Unexpected situation: a new-line was encountered which doesn't mark the end of a subtitle. Occurred when processing line " + str(l+1) + " in the .srt file. Continuing...")

			elif any([v == None for v in hold]): # if any value in hold is still == None
				print("Unexpected situation: a new-line was encountered, while the subtitle is incomplete. Occurred when processing line " + str(l+1) + " in the .srt file. Continuing...")

			else: # this is the expected situation; a new line marks the end of the subtitle
				subs_original.append(Subtitle(hold[0], hold[1], hold[2])) # load the hold into a new subtitle
				hold = [None, None, None] # empty the hold

		elif line.isnumeric(): # if the line only contains a number. this represents the index of the subtitle (subtitle #), and marks the beginning of a subtitle.
			if any([v != None for v in hold]): # if the hold isn't empty
				if [v == None for v in hold] == [False, False, True]: # if the hold contains an index and timerange, but no text. this situation is uncommon but can happen if a subtitles text just happens to be a number.
					print("Uncommon situation: it appears that a subtitles' text happens to just be a number. Encountered on line "  + str(l+1) + " in the .srt file. Continuing...")

				else:
					print("Error: a purely numerical line was encountered on line "  + str(l+1) + " in the .srt file (no, it doesn't appear to be the subtitle text by coincidence). Exiting...")
					raise SystemExit

			else:
				hold[0] = line

		elif "-->" in line: # this string is found in the time-range in .srt files
			if [v != None for v in hold] == [True, False, False]:
				hold[1] = line

			else:
				print("Error: a time-range was encountered where it wasn't expected, on line "  + str(l+1) + " in the .srt file. Exiting...")
				raise SystemExit

		else: # it is the subtitle text
			if [v != None for v in hold] == [True, True, False]: # => first/only line of subtitle text
				hold[2] = line

			elif all([v != None for v in hold]): # => the subtitle has more than one line
				hold[2] += (" " + line)

			else:
				print("Error: subtitle text was encountered where it wasn't expected, on line "  + str(l+1) + " in the .srt file. Exiting...")
				raise SystemExit

	if all([v != None for v in hold]): subs_original.append(Subtitle(hold[0], hold[1], hold[2]))
	elif any([v != None for v in hold]):
		print("The last subtitle is incomplete. Exiting...")
		raise SystemExit

	if cfg.verbosity >= 5:
		for x in subs_original: print(x)
	# }

	# translate each subtitle {
	print("Translating subtitles... \r", end="")
	for i, sub in enumerate(subs_original):
		subs_translated.append(Subtitle(sub.number, sub.timeRange, Translate.translateText_robust(sub.text, cfg.inLang, cfg.outLang[0])))

		# if ((i % (len(subs_original)/10)) == 0) or ((i % (len(subs_original)/10)) < ((i-1) % (len(subs_original)/10))):
		# 	print(".", end="")
		print("Translating subtitles... " + str((i+1) / len(subs_original)*100) + "% complete\r", end=""),
	print()

	if cfg.testingMode:
		for x in subs_translated: print(x)
	# }

	# write results into a new .srt file {
	print("Writing results to new file...", end="")
	with io.open(cfg.outPath + G.basename(filename) + " -" + cfg.outLang[0] + G.extension(filename), mode="w", encoding="utf-8") as file_translate: # output will be in utf-8 no matter the input .srt encoding. i did this because google translate api outputs in utf-8.
		for i, sub in enumerate(subs_translated):
			if i == len(subs_translated)-1:
				file_translate.write(sub.number + "\n" + sub.timeRange + "\n" + sub.text)
			else: file_translate.write(str(sub) + "\n")
	print(" done")
	# }

for srtFile in fileList.input_srt: _translateSubtitles(srtFile)

print("End of script.")

# SOURCES ==========================================
	
	# translatePDF.py
	# Download Links in order (aria2c).py