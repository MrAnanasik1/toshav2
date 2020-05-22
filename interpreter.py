import requests
import json
import logging
logging.basicConfig(format="%(levelname)s %(asctime)s - %(message)s")
from typing import Text, List, Dict, Any, Union, Optional, Tuple

DEF_URL = "https://ai.twin24.ai/api/v1"
DEF_TOKEN = "a49ed8b79f05c6caaa94a79acc586c5fb4027f53" # токен для пользователя masha16

DEF_AGENT_UUID = "398a5424-faf5-4aeb-b50d-51835c037970" # Tosha_v2
DEF_TIMEZONE = "UTC"

class TwinAPI:
    """
    Класс реализует запросы к twin
    """
    def __init__(self, url = DEF_URL,
                 token = DEF_TOKEN):
        """

        :param url:
        :param token:
        """
        self.url = url
        self.token = token

    def rasa_nlu_parse(self,agent_uuid: str, timezone: str, query: str) ->Dict:
        """
        Парсинг текстового запроса
        :param agent_uuid: uuid агента
        :param timezone: временнАя зона
        :param query: фраза
        :return:  json-encoded content, словарь
        """
        data = {
            "agent_uuid": agent_uuid,
            "timezone": timezone,
            "query": query
        }
        headers = {
            "Accept": "application/json",
            "Authorization": "Token {token}".format(token=self.token)
        }
        r = requests.post("{url}/rasa_nlu/parse/".format(url=self.url),
                          data=data,
                          headers=headers)
        return r.json()

    def intent_types(self,agent_uuid: str) -> List[Dict]:
        """
        Получить все интенты для агента
        :param agent_uuid: uuid агента
        :return:
        """
        headers = {
            "Accept": "application/json",
            "Authorization": "Token {token}".format(token=self.token)
        }
        r = requests.get("{url}/agents/{agent_uuid}/intent_types/".format(url=self.url, agent_uuid=agent_uuid), headers=headers)
        return r.json()


class TwinInterpreter:
    """
    Интерпретатор фраз TWIN
    """
    def __init__(self,
                 api = TwinAPI(),
                 agent_uuid: str = DEF_AGENT_UUID,
                 timezone: str = DEF_TIMEZONE):
        self.api = api
        self.agent_uuid = agent_uuid
        self.timezone = timezone


    def parse(self, text: str)-> Tuple[Dict,List[Dict]]:
        """

        :param text:
        :return:
        """
        default_dict = {'confidence': 0.9999, 'name': 'default'}
        try:
            r = self.api.rasa_nlu_parse(self.agent_uuid, self.timezone, text)
            intent = r.get("intent",default_dict)
            entities = r.get("entities",list())
            return intent, entities
        except Exception as e:
            logging.error(str(e))
            return (default_dict, list())

if __name__ == "__main__":
    # create twin api
    api = TwinAPI(url = DEF_URL, token=DEF_TOKEN)
    # получить все намерения, зарегистрированные для агента
    intent_types = api.intent_types(DEF_AGENT_UUID)
    # create twin интерпретатор
    twin = TwinInterpreter(api=api,agent_uuid=DEF_AGENT_UUID,timezone=DEF_TIMEZONE)
    # распарсить запрос
    intent, entities = twin.parse("Привет Тоша, меня зовут Татьяна Ивановна")
    print(intent)
    pass