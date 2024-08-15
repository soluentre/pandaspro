import re
from typing import Any, List, Union


def wildcardread(varkey, stringlist):
    """
    This is the wildcard reader function which can parse containing-wildcard varnames into meaningful list of varnames
    For example: mak* can return the list of ["make1", "make2", "make3"] which can be further used to slice dataframes

    :param varkey: a variable key with wildcard in it to match one or more variables
    :param stringlist: a list of vars with wildcards
    :return:
    """
    if '--' in varkey:
        crange = re.split(r'\s*--\s*', varkey)
        element1 = crange[0]
        element2 = crange[1]
        if element1 not in stringlist or element2 not in stringlist:
            print('Invalid column name')
            return None
        if list(stringlist).index(element1) > list(stringlist).index(element2):
            element1, element2 = element2, element1
        return list(stringlist)[list(stringlist).index(element1): list(stringlist).index(element2) + 1]

    else:
        varkey = varkey.replace('*', '<WILDCARD_STAR>').replace('?', '<WILDCARD_QMARK>')
        pattern = re.escape(varkey)
        pattern = pattern.replace('<WILDCARD_STAR>', '.*').replace('<WILDCARD_QMARK>', '.')
        # noinspection PyTypeChecker
        pattern = '^' + pattern + '$'
        regex = re.compile(pattern, re.DOTALL)
        matching_strings = [s for s in stringlist if regex.match(s)]
        return matching_strings


def str2list(inputstring: str) -> Union[List[str], List[Union[str, Any]]]:
    """
    This function is used to turn a string of vars to a list object
    Python can not automatically parse list of vars as written in a string separated by space, like "make price mpg rep78" as comparing to Stata
    And this function will serve as the parser to separate the string with spaces into var/var with wildcard sections

    :param inputstring: the key input a string with many varnames separated by X number of spaces
    Note: you can use three types of wildcard: * ? -, as supported with the wildcardread function

    :return: a list of varnames
    """
    pattern = r'\w+\s*--\s*\w+'
    match = re.findall(pattern, inputstring)
    if not match:
        newlist = [s.strip() for s in inputstring.split(';')]
    else:
        for index, item in enumerate(match):
            inputstring = inputstring.replace(item, '__' + str(index) + '__')
        aloneitem = inputstring.split(';')
        for index, item in enumerate(match):
            newlist = [item if s.strip() == '__' + str(index) + '__' else s.strip() for s in aloneitem]

    # noinspection PyUnboundLocalVariable
    return newlist


def parse_wild(promptstring: str, checklist: list, dictmap: dict = None):
    """
    This function will return the searched varnames from a python dataframe according to the prompt string

    :param checklist: list
    :param promptstring: for example: "name* title*", must separate by blanks, meaning names should not contain blanks
    :param dictmap: dictionary to convert abbr names

    :return: a list of available varnames
    """
    varlist = []
    result_list = []
    for varkey in str2list(promptstring):
        if dictmap and varkey in dictmap.keys():
            varkey = varkey.lower()
            for term in dictmap[varkey]:
                # -- debug tests
                # print(wildcardread(checklist, term))
                varlist += wildcardread(term, checklist)

        else:
            # -- debug tests
            # print(wildcardread(checklist, varkey))
            varlist += wildcardread(varkey, checklist)
    for x in varlist:
        if x not in result_list:
            result_list.append(x)
    return result_list


def clean_keys(input_dict):
    return {re.sub(r'[^a-zA-Z0-9]', '', key): value for key, value in input_dict.items()}


def clean_string(input_string):
    return re.sub(r'[^a-zA-Z0-9]', '', input_string)


def encapsulate_lists(module):
    lists_dict = {}
    for name, value in vars(module).items():
        if isinstance(value, list):
            lists_dict[name] = value
    return lists_dict


def parse_method(input_string):
    """
    Parses the given string to extract the method name. If parameters are present, it also extracts them into a dictionary.
    Includes internal helper functions to parse values and intelligently split parameter strings.
    """

    import ast

    def parse_value(value_local):
        """
        Attempts to convert a string value to its corresponding data type.
        """
        try:
            # Try to parse it as a complex data type (list, tuple, dict)
            my_parsed_value = ast.literal_eval(value_local)
            # If it's a list with elements, ensure each element is a string
            if isinstance(my_parsed_value, list):
                return [str(element).strip() for element in my_parsed_value]
            return my_parsed_value
        except (ValueError, SyntaxError):
            # Handle non-literal lists, assuming they are lists of strings
            if value_local.startswith('[') and value_local.endswith(']'):
                # Remove the brackets and split by commas not enclosed in brackets
                elements = smart_split_params(value_local[1:-1])
                return [element.strip() for element in elements]
            else:
                # If ast.literal_eval fails, fall back to trying int, then float, then return as string
                try:
                    return int(value_local)
                except ValueError:
                    try:
                        return float(value_local)
                    except ValueError:
                        # Return as string if it's neither a number nor a complex type
                        return value_local

    def smart_split_params(params_string_local):
        """
        Intelligently splits the parameters string, taking into account commas
        within lists, tuples, and dictionaries.
        """
        params = []
        bracket_level = 0  # Tracks the level of nested brackets
        current_param = ''

        for char in params_string_local:
            if char in '([{':
                bracket_level += 1
            elif char in ')]}':
                bracket_level -= 1
            elif char == ',' and bracket_level == 0:
                # Only split at top-level commas
                params.append(current_param.strip())
                current_param = ''
                continue

            current_param += char

        # Add the last parameter
        if current_param:
            params.append(current_param.strip())

        return params

    # Check if the input string contains parentheses
    if '(' in input_string and ')' in input_string:
        # Extract method name and parameters if parentheses are present
        method_pattern = r'^(.*?)\((.*)\)$'
        match = re.match(method_pattern, input_string)

        if match:
            method_name = match.group(1)  # Extract method name
            params_string = match.group(2)  # Extract parameters string

            # Use smart_split_params to handle complex parameter values
            params_list = smart_split_params(params_string)
            params_dict = {}

            # Convert each parameter to a key-value pair in the dictionary
            for param in params_list:
                key, value = param.split('=')
                parsed_value = parse_value(value.strip())  # Parse value to correct type
                params_dict[key.strip()] = parsed_value

            return method_name, params_dict
    else:
        # If there are no parentheses, return only the method name
        return input_string, {}
