'''工具'''
from email.header import decode_header
import time
import re
from html import unescape

# 共通方法类
class BaseCommon :
    @classmethod
    # 替换特殊字符
    def rpl(self,str1):
        str2 = eval(repr(str1).replace('\\\\t', '\\t'))
        str3 = str2.replace("\\'", "'")
        return str3

    # 忽略charset
    def decode_str(self,s):
        value, charset = decode_header(s)[0]
        if charset:
            value = value.decode(charset, 'ignore')
        return value

    # 时间格式转化
    def trans_format(self,time_string, from_format, to_format='%Y.%m.%d %H:%M:%S'):
        """
        @note 时间格式转化
        :param time_string:
        :param from_format:
        :param to_format:
        :return:
        """
        time_struct = time.strptime(time_string, from_format)
        times = time.strftime(to_format, time_struct)
        return times

    # 获取编码格式
    def guess_charset(self,msg):
        charset = msg.get_charset()
        if charset is None:
            content_type = msg.get('Content-Type', '').lower()
            pos = content_type.find('charset=')
            if pos >= 0:
                charset = content_type[pos + 8:].strip()
        if (charset == '"gb2312"'):
            charset = '"gb18030"'
        return charset

    # 替换html标签元素
    def html_to_plain_text(self,html):
        text = re.sub('<head.*?>.*?</head>', '', html, flags=re.M | re.S | re.I)
        text = re.sub('<a\s.*?>', ' HYPERLINK ', text, flags=re.M | re.S | re.I)
        text = re.sub('<.*?>', '', text, flags=re.M | re.S)
        text = re.sub(r'(\s*\n)+', '\n', text, flags=re.M | re.S)
        return unescape(text)