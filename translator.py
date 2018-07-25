from googletrans import Translator
import n2w

'''
Small module to handle integer to word conversion,
as well as word translation to the needed language (portuguese in this case)
'''

translator = Translator()

def translate(num):
    number = n2w.convert(num) #Converts integer to its equivalent word (e.g.: 1 -> 'one')
    return translator.translate(number, dest='pt').text #Translates the previously created word into portuguese
