#!usr/bin/env python3

__author__ = 'Myeongseon Choi key262yek@gmail.com'
__date__ = "2017/12/11"
__version__ = '1.1.0'
__version_info__ = (1,1,0) #Provide removing files including. previous error.


'''================================= Import Module ==============================='''

import os, logging,platform
import stat #Change permission.
import win32security #unresolved , you can get it from pypiwin32
import ntsecuritycon as con #unresolved
import re #Regex
import chardet #Unresolved import


'''================================ Parameter declaration ========================='''

default_language = 'ko-KR'


'''================================= Class ==============================='''

class smiItem() :
	def __init__(self,path,head,body) :
		self.path = path #Path of smi file. It need for exporting srt file.
		self.head = head #Head of smi
		self.body = body  #Body of smi file

		language_info = self.find_language()
		# structure : CC name : (srt_subtitles, language name)
		# ex) 'ENCC' : ([...], 'en-US')
		# converted srt files for each languages
		self.converted = {name : [[],lang] for name,lang in language_info}
		logging.info('Declare smi files is done. Directory : {}'.format(self.path))

	def find_language(self) :
		#Find languages of subtitles and its local name.
		#ex) .ENCC {Name: 'English Captions'; lang: en-US; SAMIType : CC;}
		#space after period and ENCC, space and text between {}
		language_setting = re.findall('[ ]*\.([^ ]*)[ ]*\{([^}]*)\}',self.head)
		language_info = []
		for name,setting in language_setting :
			#Remove space.
			nospace_setting = re.sub(r'\s+','',setting)

			#split by ';' last term must be blank(i.e. ''). we need to remove it also.
			split_setting = re.split(';|:',nospace_setting)[:-1]

			#make dictionary from it
			dict_setting = {category : split_setting[2*idx+1] for idx,category in enumerate(split_setting[::2])}
			if 'lang' in dict_setting :
				lang = dict_setting['lang']
			else :
				lang = default_language #In my case, 'ko-KR' is default.
				logging.warning('the language entity is empty in the head of the file {}'.format(self.path))
			language_info.append((name,lang))
		logging.debug('find_language finished')
		return language_info

	def remove_meaningless_tags(self) :
		#Remove tags which are not supported for srt subtitle format.
		meaningless = ['b','i','u','basefont','bdo','big','blockquote','caption','center','col','colgroup','dd','dl','dt','div','h1','h2','h3','h4','h5','h6','hr','img','li','ol','p','pre','q','s','small','span','strike','sub','sup','table','tbody','td','th','tr','tt','tfoot','thead','ul']
		for key in meaningless :
			key_hat = r'</?'+re.escape(key)+r'>'
			self.body = re.sub(key_hat,'',self.body,flags=re.IGNORECASE)

		#Remove <P Class = ..CC ID = ... > tag
		self.body = re.sub(r'<[ ]*[Pp] [Cc]lass[ =]*[A-Za-z]*[Cc]{2}[ ]*[IDid]{2}[^>]*>[^<]*','',self.body)

		#Remove <font color .. > tag
		self.body = re.sub(r'<[ /]*font[^>]*>','',self.body,flags=re.IGNORECASE)

		#Remove <! ... --> tag. those are comment format of sami.
		self.body = re.sub(r'<![-]{,2}(.*?)[-]{1,2}>','',self.body)

		#Logging
		logging.debug('remove_meaningless_tags finished')

	'''
	def replace_br_tags(self) :
		#Change <br> tags to '\n'
		self.body = re.sub(r'<br>','\n',self.body,flags = re.IGNORECASE)
		logging.debug('change_br_tags finished')
	'''

	def convert_subtitles(self) :
		#Find all text and save in proper list of subtitles.
		end_point,smi_starttime = 0,0
		for synctag in re.finditer(r'<sync\s*start[ =]*(\d*)>',self.body,re.IGNORECASE) : # Find the tag <STNC Start = (time)>
			'''Matching algorithm.
			<sync start = (start time)>(end_point)
			.
			.
			.
			(start_point)<sync start = (end time)>

			search item between end_point of previous tag and start_point of current tag.
			This algorithm may ignore the last subtitle which is not wrapped by <sync start> tags, however for convenient I use this method.
			'''
			#start_point = index of the first character '<' of <sync..
			#smitime = time information of subtitle
			#end point of previous tag is needed to search P Class tag. Hence replace it in the end of current iteration.
			start_point = synctag.start()
			smitime = int(synctag.group(1))

			#It this tag is the first tag of the group of subtitles, move to next tag.
			if smi_starttime == 0 or smitime < smi_starttime :
				start_point = synctag.start()
				end_point = synctag.end()
				smi_starttime = smitime
				continue

			smi_endtime = smitime
			srt_starttime = srt_time(smi_starttime)
			srt_endtime = srt_time(smi_endtime)
			#Read <P Class = nametag>
			#There can be multiple P Class tags in a single Sync tag. (((?!<[ ]*P).)*) makes it works including such cases.
			for langtag in re.finditer(r'<[ ]*[Pp][ ]*[Cc]lass[ =]*([A-Za-z]*CC)[ ]*>(((?!<[ ]*P).)*)',self.body[end_point:start_point]) :
				name = langtag.group(1)
				msg = langtag.group(2)
				if '&nbsp' in msg :
					continue
				if not name in self.converted :
					#Sometimes, there are some tricky files which do not declare CC names in the head. Moreover the CC name in body is 'UnknownC'
					self.converted[name] = ([],default_language) #It may be more probable the unknowncc are the default lanugage.
					logging.warning('the unknowncc appear in the file {}'.format(self.path))

				subtitles,lang = self.converted[name]
				subtitle = srt_format(srt_starttime,srt_endtime,msg)
				subtitles.append(subtitle)

			smi_starttime = smitime
			end_point = synctag.end()
		logging.debug('convert_subtitles is finished')

	def write_srt(self) :
		#write srt file from subtitles
		filename = self.path[:-3] #remove 'smi' from the original smi file.
		for subtitles,lang in self.converted.values() :
			if not subtitles : #There exist the smi files which do not include subtitles of language declared in the head.
				continue
			srt_output = open(filename+lang[:2]+'.srt','wb')
			for idx,subtitle in enumerate(subtitles) :
				subtitle_with_number = '{}\r\n'.format(idx+1) + subtitle

				srt_output.write(subtitle_with_number.encode('utf-8'))
			srt_output.close()
		logging.info('Converted finished. file directory : {}'.format(self.path))


''' ====================================== Fucntions ============================================'''


#Read contents from smi files.
def read_smi(smi_path) :
	#Read smi file.
	smi = open(smi_path,'rb')
	contents = smi.read()
	smi.close()

	#Check the encoding of contents
	encoding = chardet.detect(contents)['encoding']

	#Raise error when there is no encoding detected.
	if encoding == None :
		logging.warning('Cannot find a proper encoding of file {}'.format(smi_path))
		return None

	#Re-encode contents into unicode
	contents = contents.decode(encoding)

	#Remove newline character. It is need for regex multiline search
	contents = re.sub(r'[\n\r\t]+','',contents)

	#cut head and body by <HEAD> and <BODY> tags.
	head_contents = re.search('<HEAD>(.*)</HEAD>',contents,re.IGNORECASE)
	body_contents = re.search('<BODY>(.*)</BODY>',contents,re.IGNORECASE)

	if not (head_contents and body_contents):
		logging.warning('Head or Body is missing in the file {}'.format(smi_path))
		return None

	#Declare and return smi class
	logging.debug('read_smi finished for the file : {}'.format(smi_path))
	return smiItem(smi_path,head_contents.group(1),body_contents.group(1))

def srt_time(smitime) :
	#smi time = 0000000 (only ms)
	#srt time = hours:minutes:seconds,ms
	#Convert input smi time to srt time format.
	time,ms = divmod(smitime,1000)
	time,sec = divmod(time,60)
	hour,minute = divmod(time,60)
	logging.debug('srt_time finished')
	return '{}:{}:{},{}'.format(hour,minute,sec,ms)


def srt_format(srt_starttime,srt_endtime,msg) :
	#Make srt subtitle from given times and msg.
	#subtitle number will be added while writing.
	msg = msg.replace('<br>','\r\n')
	logging.debug('srt_format finished')
	return '{} --> {}\r\n{}\r\n\r\n'.format(srt_starttime,srt_endtime,msg)

def remove_empty_srt(srt_path) :
	#Error correcting.
	#Becuase of failure of algorithm, a lot of empty srt files are made in various directory.
	srt_file = open(srt_path,'rb')
	if not srt_file :
		return
	text = srt_file.read()
	if not text :
		if platform.system() is 'Windows' :
			userx, domain, type = win32security.LookupAccountName("","Everyone")
			sd = win32security.GetFileSecurity(srt_path,win32security.DACL_SECURITY_INFORMATION)
			dacl = sd.GetSecurityDescriptorDacl()
			dacl.AddAccessAllowedAce(win32security.ACL_REVISION	, con.FILE_ALL_ACCESS, userx)
			sd.SetSecurityDescriptorDacl(1,dacl,0)
			win32security.SetFileSecurity(srt_path,win32security.DACL_SECURITY_INFORMATION,sd)

		else :
			os.chmod(srt_path, stat.S_IWOTH)
		os.remove(srt_path)
		logging.info('The empty srt file is found in directory {}'.format(srt_path))
		return

''' =================================  Main ==================================='''


#Get directories of smi files in the subdirectories of current working directory
cwd = os.getcwd()
logging.basicConfig(filename = 'smi2srt.log',level = logging.WARNING)

logging.debug('program start')
smifiles = []
for current_directory,subdirectories,files in os.walk(cwd) :
	for name in files : #Check the name of files.
		if name.endswith('.srt') : #Check empty srt files.
			remove_empty_srt(os.path.join(current_directory,name))
		elif name.endswith(".smi") : #and not name[:-3]+'ko.srt' in files :
			#Non converted files only. (However you should check that the already existed file could be empty because of previous error.)
			smifiles.append(os.path.join(current_directory,name))
logging.debug('file directory reading is end')


for smi_path in smifiles :
	smi_contents = read_smi(smi_path)
	if not smi_contents :
		continue
	smi_contents.remove_meaningless_tags()
	smi_contents.convert_subtitles()
	smi_contents.write_srt()

logging.debug('whole program finished well.')



