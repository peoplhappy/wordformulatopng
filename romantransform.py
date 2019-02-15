def transform_alabo2_roman_num_upper(one_num):
    '''
    将阿拉伯数字转化为罗马数字
    '''
    num_list = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
    str_list = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"]
    res = ''
    for i in range(len(num_list)):
        while one_num >= num_list[i]:
            one_num -= num_list[i]
            res += str_list[i]
    return res

def transform_alabo2_roman_num_lower(one_num):
    '''
    将阿拉伯数字转化为罗马数字
    '''
    num_list = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
    str_list = ["m", "cm", "d", "cd", "c", "xc", "l", "xl", "x", "ix", "v", "iv", "i"]
    res = ''
    for i in range(len(num_list)):
        while one_num >= num_list[i]:
            one_num -= num_list[i]
            res += str_list[i]
    return res

def transform_roman_num2_alabo(one_str):
    '''
    将罗马数字转化为阿拉伯数字
    '''
    define_dict = {'I': 1, 'V': 5, 'X': 10, 'L': 50, 'C': 100, 'D': 500, 'M': 1000}
    if one_str == '0':
        return 0
    else:
        res = 0
        for i in range(0, len(one_str)):
            if i == 0 or define_dict[one_str[i]] <= define_dict[one_str[i - 1]]:
                res += define_dict[one_str[i]]
            else:
                res += define_dict[one_str[i]] - 2 * define_dict[one_str[i - 1]]
        return res
