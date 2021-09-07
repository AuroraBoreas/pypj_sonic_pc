"""
this is a simple class to translate long texts using free Translator(F0) service, which is provided by MS Azure.

it has the following functionalities:
- translates from A language to B language
- displays translation result
- lang list: [https://docs.microsoft.com/en-us/azure/cognitive-services/translator/language-support]

Changelog
- v0.0.1, initial version

@ZL, 20210121

"""

import requests, uuid, json

class AzureTranslator:
    # Add your subscription key and endpoint
    __subscription_key = "399a485f234544dcb07d2723ff65186a" # key may expire. have to freshen it from your subscription
    __endpoint = "https://api.cognitive.microsofttranslator.com/"

    # Add your location, also known as region. The default is global.
    # This is required if using a Cognitive Services resource.
    __location = "japaneast"

    __path = '/translate'
    __constructed_url = __endpoint + __path

    def __init__(self, raw_text: str, from_lang_code: str = 'ja', to_lang_code: str = 'zh-Hans'):
        self.params = {
            'api-version': '3.0',
            'from': 'ja',   # from lang code
            'to': 'zh-Hans' # to lang code
        }
        self.constructed_url = self.__endpoint + self.__path

        self.headers = {
            'Ocp-Apim-Subscription-Key': self.__subscription_key,
            'Ocp-Apim-Subscription-Region': self.__location,
            'Content-type': 'application/json',
            'X-ClientTraceId': str(uuid.uuid4())
        }

        # You can pass more than one object in body.
        self.body = [{
            'text': raw_text
        }]

    def start(self):
        request = requests.post(self.__constructed_url, params=self.params, headers=self.headers, json=self.body)
        response = request.json()
        # res = json.dumps(response, sort_keys=True, ensure_ascii=False, indent=4, separators=(',', ': '))
        res = response[0]['translations'][0]['text']
        return res