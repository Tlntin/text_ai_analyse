import PySimpleGUI as sg
from aip import AipNlp
from docx import Document
import os
import pickle
import random
import time


class TextAIAnalyse(object):
    """
    :param
    """

    def __init__(self, doc_path, app_id, api_key, secret_key):
        """
        :param doc_path：文章路径
        :param app_id: 应用id，自己去百度ai控制台构建一个应用，就会有id了
        :param api_key:
        :param secret_key:
        """
        self.client = AipNlp(app_id, api_key, secret_key)
        self.document = Document(doc_path)
        self.doc_path = doc_path
        text_list1 = self.filter_style()
        text_list2 = self.filter_short_text(text_list1, 12)
        self.text_list3 = self.split_text(text_list2)

    def filter_style(self):
        """
        样式过滤
        :param
        """
        delete_style = ['Title', 'Heading 1', 'Quote']  # 去除标题，一级标题，图表链接
        list1 = [x.text for x in self.document.paragraphs if x.style.name not in delete_style]
        return list1

    @staticmethod
    def filter_short_text(list1: list, length: int):
        """
        去除短文本
        :param list1: 列表格式的文本集
        :param length:最短文本长度
        """
        list2 = [x for x in list1 if len(x) > length]
        return list2

    def split_text(self, list1: list):
        """
        对段落进行分句，粗糙分词
        :param
        """
        list2 = []
        for x in list1:
            x_list1 = x.split('。')  # 以句号进行分词,分号暂时不考虑
            for xx in x_list1:
                if xx[-1:] not in ['。', '；', '：']:  # 如果本局不是以句号结尾，则给它加上句号
                    xx += '。'
                    list2.append(xx)
        list3 = self.filter_short_text(list2, 10)
        return list3

    def split_text2(self, list1: list):
        """
        对段落进行分句，加上分号
        :param
        """
        list2 = []
        for x in list1:
            x_list1 = x.split('。')  # 以句号进行分词
            for xx in x_list1:
                x_list2 = xx.split('；')  # 以中文分号进行分词
                for xxx in x_list2:
                    if xxx[-1:] not in ['。', '；', '：']:  # 如果本局不是以句号结尾，则给它加上句号
                        xxx += '。'
                        list2.append(xxx)
        list3 = self.filter_short_text(list2, 10)
        return list3

    def ai_analyse(self, text1):
        """
        AI对句子进行纠错
        :param
        """
        result1 = None
        try:
            result1 = self.client.ecnet(text1)
        except Exception as err:
            return False
        vet = result1['item']['vec_fragment']  # 可替换词
        score = result1['item']['score']  # 评分
        if len(vet) == 0:
            return False  # 没有错误
        elif score > 0.5:  # 如果可信度
            return result1  # 返回分析

    def save_analyse(self, dict1):
        """
        :param
        """
        base_name = os.path.basename(self.doc_path)  # 文件名
        if not os.path.exists('./分析结果'):
            os.mkdir('./分析结果')
        path = './分析结果/' + base_name[:8] + '.csv'
        text = dict1['text']  # 错误的句子
        correct_query = dict1['item']['correct_query']  # 正确的句子
        vec_list = dict1['item']['vec_fragment']  # 分析过程
        vec_error = ','.join([x['ori_frag'] for x in vec_list])  # 提取错误的句子
        vec_true = ','.join([x['correct_frag'] for x in vec_list])  # 提取正确的句子
        score = dict1['item']['score']  # 评分
        str2 = '{},{},{},{},{}\n'.format(text, correct_query, vec_error, vec_true, score)
        if not os.path.exists(path):  # 如果路径不存在
            f = open(path, 'wt', encoding='utf-8-sig')
            str1 = '错误的句子, 正确的句子,错误的词为,正确的词为, 可信度评分\n'
            f.write(str1)
            f.write(str2)
            f.close()
        else:
            f = open(path, 'at', encoding='utf-8-sig')
            f.write(str2)
            f.close()


if __name__ == '__main__':
    # 菜单栏
    menu_def = [
        ['&菜单', ['使用说明', '更新记录']],
        ['&文件', ['载入配置', '保存配置']]
    ]
    # 布局栏
    layout1 = [
        [sg.Menu(menu_def, tearoff=True)],
        [sg.Text('APP_ID:', size=(12, None)), sg.Input(key='app_id')],
        [sg.Text('API_KEY', size=(12, None)), sg.Input(key='api_key')],
        [sg.Text('SECRET_KEY', size=(12, None)), sg.Input(key='secret_key')],
        [sg.Text('文件位置'), sg.Input(key='file_name', size=(51, None))],
        [sg.FileBrowse('选择文件', target='file_name'), sg.Button('开始检测'), sg.Button('退出')]
    ]
    # 窗口栏
    windows1 = sg.Window('小错误检测器V1.0', layout=layout1)
    for i in range(10):
        event1, value1 = windows1.read()
        if event1 in ('退出', None):
            break
        elif event1 == '使用说明':
            sg.popup('1.搜索百度AI开放平台', '2.点击控制台，注册并登录', '3.选择自然语音处理', '4.创建应用',
                     '5.填写appid, api_key, secret_key', '6.选择需要纠错的文件', '7.点击开始检测', title='使用说明')
        elif event1 == '更新记录':
            sg.popup('暂时没有更新记录', title='提示')
        elif event1 == '保存配置':
            APP_ID = windows1['app_id'].get()
            API_KEY = windows1['api_key'].get()
            SECRET_KEY = windows1['secret_key'].get()
            file_path = windows1['file_name'].get()
            if len(APP_ID) > 3 and len(API_KEY) > 5 and len(SECRET_KEY) > 5:
                dict1 = {
                    'app_id': APP_ID,
                    'api_key': API_KEY,
                    'secret_key': SECRET_KEY
                }
                with open('info.pkl', 'wb') as f:
                    pickle.dump(dict1, f)
                sg.popup('保存完毕', '已经生成一个info.pkl文件到本地', title='提示', auto_close=True,
                         auto_close_duration=3)
            else:
                sg.popup('请检查你的api相关信息是否填写完成', title='错误提示')
        elif event1 == '载入配置':
            if not os.path.exists('info.pkl'):
                sg.popup('没有找到你的配置文件info.pkl', '请检查你的文件是否在当前路径', title='错误提示')
            else:
                with open('info.pkl', 'rb') as f:
                    dict2 = pickle.load(f)
                    windows1['app_id'].update(dict2['app_id'])
                    windows1['api_key'].update(dict2['api_key'])
                    windows1['secret_key'].update(dict2['secret_key'])
                    sg.popup('配置文件载入完毕', title='提示', auto_close_duration=3, auto_close=True)
        elif event1 == '开始检测':
            APP_ID = windows1['app_id'].get()
            API_KEY = windows1['api_key'].get()
            SECRET_KEY = windows1['secret_key'].get()
            file_path = windows1['file_name'].get()
            if len(APP_ID) > 3 and len(API_KEY) > 5 and len(SECRET_KEY) > 5:
                doc = TextAIAnalyse(file_path, APP_ID, API_KEY, SECRET_KEY)
                text_list3 = doc.text_list3
                sg.popup('开始检测，共有{}句'.format(len(text_list3)),
                         '预计用时{}秒'.format(len(text_list3)*2), auto_close_duration=5, auto_close=True)
                layout2 = [
                    [sg.Text('处理进度条'), sg.ProgressBar(len(text_list3), orientation='h', key='bar', size=(50, 20))],
                    [sg.Button('取消')]
                ]
                windows2 = sg.Window(title='进度条', layout=layout2)
                bar = windows2['bar']
                for ii in range(len(text_list3)):
                    event2, value2 = windows2.read(timeout=10)
                    if event2 in ('取消', None):
                        break
                    result2 = doc.ai_analyse(text_list3[ii])
                    if result2:
                        print(result2)
                        doc.save_analyse(result2)
                    time.sleep(0.5 + random.random() / 10)
                    bar.UpdateBar(ii + i)
                windows2.close()
                sg.popup('已经检测完成', '并且生成了一个“分析结果”文件夹到本地', title='提示')
            elif len(file_path) < 5:
                sg.popup('亲！', '你还没有选择检测的文件', title='提示')
            else:
                sg.popup('请输入你的api信息', '详情请查看使用说明', title='提示')
    windows1.close()
