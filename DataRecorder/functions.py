from typing import Union


def _data_to_list(data: Union[list, tuple, dict],
                  before: Union[list, dict, None] = None,
                  after: Union[list, dict, None] = None) -> list:
    """将传入的数据转换为列表形式          \n
    :param data: 要处理的数据
    :param before: 数据前的列
    :param after: 数据后的列
    :return: 转变成列表方式的数据
    """
    return_list = []
    if not isinstance(data, (list, tuple, dict)):
        data = [data]
    if not isinstance(before, (list, tuple, dict)):
        data = [before]
    if not isinstance(after, (list, tuple, dict)):
        data = [after]

    for i in (before, data, after):
        if isinstance(i, dict):
            return_list.extend(list(i.values()))
        elif i is None:
            pass
        elif isinstance(i, list):
            return_list.extend(i)
        elif isinstance(i, tuple):
            return_list.extend(list(i))
        else:
            return_list.extend([str(i)])

    return return_list


def _data_to_list_or_dict(data: Union[list, tuple, dict],
                          before: Union[list, tuple, dict, None] = None,
                          after: Union[list, tuple, dict, None] = None) -> Union[list, dict]:
    """将传入的数据转换为列表或字典形式，用于记录到txt或json          \n
    :param data: 要处理的数据
    :param before: 数据前的列
    :param after: 数据后的列
    :return: 转变成列表或字典形式的数据
    """
    if isinstance(data, (list, tuple)):
        return _data_to_list(data, before, after)

    elif isinstance(data, dict):
        if isinstance(before, dict):
            data = {**before, **data}

        if isinstance(after, dict):
            data = {**data, **after}

        return data
