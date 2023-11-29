from pathlib import Path
from typing import Union


def json_read(path: Union[Path, str]) -> Union[dict, list]:
    import json

    with open(str(path), 'r', encoding='utf-8') as fp:
        data = json.load(fp)
    return data


# ? tested
def json_write(path: Union[Path, str], data: Union[dict, list]) -> None:
    import json

    with open(str(path), 'w', encoding='utf-8') as fp:
        json.dump(data, fp, ensure_ascii=False)
