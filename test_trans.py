from deep_translator import GoogleTranslator
translator = GoogleTranslator(source='auto', target='zh-TW')
text = translator.translate('부처님오신날')
print(text)
