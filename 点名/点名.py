import tkinter as tk

import xlrd

import time

import random

import datetime


class LoveYou():

    # 初始化
    def __init__(self):
        # 第1步，建立窗口window
        self.window = tk.Tk()

        # 第2步，给窗口的可视化起名字
        self.window.title('班级点名')

        # 第3步，设定窗口的大小(长＊宽)
        self.window.geometry('600x400')

        self.text = tk.StringVar()  # 创建str类型
        self.count = tk.StringVar()

        self.data = self.read_data()

        # 获取星期几
        d = datetime.datetime.now()
        self.day = d.weekday() + 1

    def read_data(self):
        '''
        数据读取
        :return:
        '''
        workbook = xlrd.open_workbook('名单.xls')

        sheet1 = workbook.sheet_by_index(0)  # sheet索引从0开始

        data = list(sheet1.col_values(0))  # 读取第一列内容


        data.pop(0)  # 把表格 1去掉
        data.pop(0)  # 把姓名 去掉

        return data

    def take(self):
        '''
        负责随机抽取同学提问
        :return:
        '''

        for s in range(10):
            '''
            后几秒慢点，制造紧张气氛
            '''
            desc = ''
            if s == 47:
                time.sleep(0.5)
            elif s == 48:
                time.sleep(0.6)
            elif s == 48:
                time.sleep(0.7)
            elif s == 49:
                time.sleep(0.9)
            else:
                time.sleep(0.1)

            classes = random.sample(self.data, 2)
            desc += "呦,你被上帝选中了:%s\n" % classes[0]
            desc += "呦,你看着也很不错呀:%s\n" % classes[1]

            if s == 49:
                self.savedesc(desc)  # 写入日志
            self.text.set(desc)  # 设置内容
            self.window.update()  # 屏幕更新




    def gettime(self):
        '''
        格式化时间
        :return:
        '''
        return time.strftime('%Y-%m-%d', time.localtime(time.time())) + "  星期" + str(self.day)


    def savedesc(self, desc):
        '''
        负责把选中的人写入到log里面
        :param desc:
        :return:
        '''
        with open('log.txt', 'a', encoding='utf-8') as f:
            f.write(self.gettime() + "\n" + desc)


    def main(self):
        '''
        主函数负责绘制
        :return:
        '''

        # 绘制日期、班级总人数等
        now = time.strftime('%Y-%m-%d', time.localtime(time.time())) + "  星期" + str(self.day)
        now += "\n班级总人数:%s人" % str(len(self.data))
        now += "\n正在合理计算中\n"

        l1 = tk.Label(self.window, fg='red', text=now, width=500, height=5)
        l1.config(font='Helvetica -%d bold' % 15)
        l1.pack()  # 安置标签

        # 绘制筛选信息
        l2 = tk.Label(self.window, fg='red', textvariable=self.text, width=500, height=3)
        l2.config(font='Helvetica -%d bold' % 30)
        l2.pack()


        # 绘制筛选按钮
        btntake = tk.Button(self.window, text="开始", width=15, height=2, command=self.take)
        btntake.pack()



        # 进入循环
        self.window.mainloop()


if __name__ == '__main__':
    loveyou = LoveYou()
    loveyou.main()

