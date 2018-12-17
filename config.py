"""
    功能:配置文件
    版本:v1.0
    作者:pwx322979
"""
import configparser
parseer=configparser.ConfigParser();
parseer.read("config.properties",encoding="utf-8")
filepath=parseer["config"]["filepath"]
cmd="Project1.exe "