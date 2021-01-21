import sys
sys.path.append('.')
import logging
logging.basicConfig(level=logging.DEBUG, format="%(asctime)s %(message)s")

from library.translator import AzureTranslator

if __name__ == "__main__":
    raw_fn = r"data\raw_texts.txt"
    new_fn = "translate\\result.txt"
    with open(raw_fn, 'r', encoding='UTF-8') as inf:
        with open(new_fn, 'w', encoding='UTF-8') as ouf:
            raw = inf.read()
            raw.replace('\n', '')
            trans = AzureTranslator(raw)
            res   = trans.start()
            ouf.writelines(res)

    logging.debug(res)