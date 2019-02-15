"""
    功能:配置文件
    版本:v1.0
    作者:pwx322979
"""
import configparser
import romantransform
parseer=configparser.ConfigParser();
parseer.read("config.properties",encoding="utf-8")
filepath=parseer["config"]["filepath"]
cmd="Project1.exe "

stylelist={
    "pydocx-list-style-type-cardinalText": {"start": 1, "increase": lambda start, index: start + index},
    "pydocx-list-style-type-decimal": {"start": 1, "increase": lambda start, index: start + index},
    "pydocx-list-style-type-decimalEnclosedCircle": {"start": 1, "increase": lambda start, index: start + index},
    "pydocx-list-style-type-decimalEnclosedFullstop": {"start": 1, "increase": lambda start, index: start + index},
     "pydocx-list-style-type-decimalEnclosedParen":{"start":1,"increase":lambda start,index: start+index},
    "pydocx-list-style-type-decimalZero":{"start":1,"increase":lambda start,index: start+index if start+index >=10 else "0"+str(start+index)},
    "pydocx-list-style-type-lowerLetter":{"start":97,"increase":lambda start,index: chr(start+index)},
    "pydocx-list-style-type-lowerRoman":{"start":1,"increase":lambda start,index: romantransform.transform_alabo2_roman_num_lower(start+index)},
    "pydocx-list-style-type-none":{"start":None},
    "pydocx-list-style-type-ordinalText": {"start": 1, "increase": lambda start, index: start + index},
    "pydocx-list-style-type-upperLetter": {"start": 65, "increase": lambda start, index: chr(start + index)},
    "pydocx-list-style-type-upperRoman": {"start": 1, "increase": lambda start,index: romantransform.transform_alabo2_roman_num_upper(start + index)}
}

