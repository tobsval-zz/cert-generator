from googletrans import Translator
import n2w

translator = Translator()

def translate(num):
    number = n2w.convert(num)
    return translator.translate(number, dest='pt').text
