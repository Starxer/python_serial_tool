import time
from tkinter import *
from tkinter import ttk, messagebox, filedialog, Spinbox
import tkinter as tk
import threading
import serial
import serial.tools.list_ports
import sys


class ComUI(object):
    ser_on = False  # 重要属性，判断串口是否启动
    to_hex = False  # 接收十六进制
    is_hex = False  # 发送十六进制
    """串口参数"""
    serial_port = ""
    serial_baudrate = 9600
    serial_bytesize = 8
    serial_parity = 'N'
    serial_stopbits = 1

    """串口对象"""
    serial_connection = None
    """串口列表"""
    serial_com, serial_info, serial_all_info = [], [], []
    """接收线程"""
    thread_recv = None

    line_break = "\r\n"
    line_end_recv = ""
    line_end_send = ""

    def __init__(self):
        """
        初始化
        """
        self.serial_com, self.serial_info, self.serial_all_info = self.get_com_list()

    """
    获取串口列表
    """

    def get_com_list(self):
        port_list = list(serial.tools.list_ports.comports())
        # print(len(port_list))

        """列表推导式代替for循环"""
        serial_com = [i[0] for i in port_list]
        serial_info = [i[1] for i in port_list]
        # serial_all_info = [i + ':' + j for i in serial_com for j in serial_info]
        serial_all_info = [serial_com[i] + ':' + serial_info[i] for i in range(len(serial_com))]
        # print('get com port success')
        return serial_com, serial_info, serial_all_info

    """
        串口开启函数
    """

    def serial_open(self):
        # 打开串口
        op = True
        print('串口参数：')
        print(self.serial_port, self.serial_baudrate, self.serial_stopbits, self.serial_bytesize, self.serial_parity)
        try:
            self.serial_connection = serial.Serial(
                port=self.serial_port,
                baudrate=self.serial_baudrate,
                parity=self.serial_parity,
                timeout=0.2,
                stopbits=self.serial_stopbits,
                bytesize=self.serial_bytesize)
        except Exception as e:
            print(e.__class__.__name__, ':', e)
            messagebox.showerror(title="错误", message="打开串口错误")
            op = False

        if op:
            self.ser_on = True
            print('串口打开成功')
            return True
        else:
            self.ser_on = False
            print('串口打开失败')
            return False

    """
        串口关闭函数
    """

    def serial_close(self):
        # 先更改ser_on发信号让子线程关闭，再关闭串口
        if self.ser_on:
            self.ser_on = False
            self.serial_connection.close()
            print('串口已关闭')
        else:
            self.ser_on = False
            print('串口未打开，无需关闭')

    """
    串口数据接受线程
    """

    def thread_recv_fun(self, text_recv):
        print('线程启动')
        while self.ser_on:  # ser_on是全局变量
            try:
                # read = self.serial_connection.readall() # readall运行太慢
                if self.serial_connection.in_waiting > 0:
                    print('读取', self.serial_connection.in_waiting)
                    read = self.serial_connection.read(self.serial_connection.in_waiting)
                    # print(__file__, sys._getframe().f_lineno, "<--",bytes(read).decode('ascii'))
                    yview = text_recv.yview()  # 这行必须在插入数据之前，否则有bug
                    text_recv.config(state=NORMAL)  # 设置为可编辑状态
                    text_recv.insert(END, self.bytes2str_or_hexstr(read, not self.to_hex) + self.line_end_recv)  #
                    text_recv.config(state=DISABLED)  # 设置为不可编辑状态
                    print(self.is_hex, self.to_hex)
                    """实现：滚动到最下方时自动滚动，往上滚动以后不再自动往下滚动"""
                    if yview[1] == 1.0:
                        text_recv.yview_moveto(1.0)
            except Exception as e:
                print(e.__class__.__name__, ':', e)
                messagebox.showerror(title="错误", message="串口连接中断，请重新连接串口")
                self.ser_on = False
                break
        print('线程关闭')

    """串口下拉框的串口刷新线程"""

    def thread_comport_update_fun(self, combo_port):
        print('线程2启动')
        while True:  # ser_on是全局变量
            if not self.ser_on:
                self.serial_com, self.serial_info, self.serial_all_info = self.get_com_list()
                combo_port.config(values=self.serial_com)
            time.sleep(1)
        # print('线程2关闭')

    """
    串口打开关闭按钮函数
    """

    def usart_ctrl(self, var, text_recv, widgets1, widgets2: tuple):
        # print(__file__,sys._getframe().f_lineno,port_,bitrate_,var.get())
        """
        若串口已启动，先关闭串口。
        适用于调整波特率的时候重启串口
        """
        if var.get() == "打开串口":

            """启动串口"""
            if self.serial_open():
                var.set("关闭串口")
                self.ser_on = True
                """启动线程"""
                self.thread_recv = threading.Thread(target=self.thread_recv_fun, args=(text_recv,))
                self.thread_recv.start()
                for i in widgets1:
                    i.config(state=NORMAL)
                for i in widgets2:
                    i.config(state=DISABLED)

        else:
            var.set("打开串口")
            self.serial_close()
            for i in widgets1:
                i.config(state=DISABLED)
            for i in widgets2:
                i.config(state=NORMAL)
            # ser.close()

    """
    保存文件按钮回调函数
    """

    def button_save_fun(self, text_recv):
        file = filedialog.asksaveasfile(defaultextension=".txt", filetypes=[("文本文件", "*.txt")])
        if file is not None:
            data = text_recv.get('1.0', tk.END)[:-1]
            data = data.replace(self.line_break, "\n")
            file.write(data)
            print("已保存到 ", file.name)
            file.close()

    """
    载入文件按钮回调函数
    """

    def button_load_fun(self, text_send):
        file = filedialog.askopenfile(defaultextension=".txt", filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")])
        if file is not None:
            text_send.delete('1.0', tk.END)
            data = file.read()
            data = data.replace("\n", self.line_break)
            text_send.insert('1.0', data)
            print("已载入 ", file.name)
            file.close()

    """
    串口发送函数
    """

    def usart_sent(self, var):
        #     print(__file__,sys._getframe().f_lineno,"-->",var)
        # print('send',var,len(var))
        var = var[:-1]
        if self.ser_on:
            data = self.str_or_hexstr2bytes(var, not self.is_hex)
            if data != None:
                self.serial_connection.write(data + self.line_end_send.encode("utf-8"))
        else:
            print("串口未开启")
        # print("-->",writedata)

    """
    串口号改变回掉函数
    """

    def combo_port_handler(self, var, var_on_off):
        """两种写法"""
        self.serial_port = var
        # self.serial_com = combo_port.get()

        # print(__file__,sys._getframe().f_lineno,var,self.serial_port)
        print(sys._getframe(0).f_code.co_name, ':', '串口切换为:', self.serial_port)

    """
    串口波特率改变回掉函数
    """

    def combo_baudrate_handler(self, varBaudrate):
        self.serial_baudrate = int(varBaudrate.get())
        print(sys._getframe(0).f_code.co_name, ':', '波特率切换为:', self.serial_baudrate)

    """
    串口数据位改变回掉函数
    """

    def combo_bytesize_handler(self, varBytesize):
        self.serial_bytesize = int(varBytesize.get())
        print(sys._getframe(0).f_code.co_name, ':', '数据位切换为:', self.serial_bytesize)

    """
    串口校验位改变回掉函数
    """

    def combo_parity_handler(self, varParity):
        parity = varParity.get()
        if parity == "NONE":
            self.serial_parity = "N"
        elif parity == "ODD":
            self.serial_parity = "O"
        elif parity == "EVEN":
            self.serial_parity = "E"
        elif parity == "MARK":
            self.serial_parity = "M"
        elif parity == "SPACE":
            self.serial_parity = "S"
        print(sys._getframe(0).f_code.co_name, ':', '校验位切换为:', self.serial_parity)

    """
    串口停止位改变回掉函数
    """

    def combo_stopbits_handler(self, varStopbits):
        self.serial_stopbits = int(varStopbits.get())
        print(sys._getframe(0).f_code.co_name, ':', '停止位切换为:', self.serial_stopbits)

    """
    点击关闭按钮时弹窗动作
    """

    def closing_procedure(self, callback, *args, **kwargs):
        response = messagebox.askyesno("退出", "确认退出")
        if response:
            """关闭串口，结束子线程"""
            try:
                self.ser_on = False
                self.serial_connection.close()
            except Exception as e:
                print(e.__class__.__name__, e)
            callback(*args, **kwargs)
        else:
            print("取消退出")

    """
    清除按钮
    """

    def text_recv_clear(self, text_recv):
        text_recv.config(state=NORMAL)  # 设置为可编辑状态
        text_recv.delete('1.0', tk.END)
        text_recv.config(state=DISABLED)  # 设置为不可编辑状态

    """
    换行复选框回调函数
    """

    def check_button_linebreak(self, var, send: bool):
        if var.get():
            if send:
                self.line_end_send = self.line_break
                print("设置发送时换行")
            else:
                self.line_end_recv = self.line_break
                print("设置接收时换行")
        else:
            if send:
                self.line_end_send = ""
                print("取消发送时换行")
            else:
                self.line_end_recv = ""
                print("取消接收时换行")

    """
    十六进制、bytes、str互转方法
    """

    def bytes2str_or_hexstr(self, val_bytes: bytes, to_str: bool):
        if to_str:
            return val_bytes.decode("utf-8")
        else:
            return val_bytes.hex()

    def str_or_hexstr2bytes(self, val: str, is_str: bool):
        if is_str:
            data = val.encode("utf-8")
        else:
            try:
                data = bytes.fromhex(val)
            except Exception as e:
                print(e.__class__.__name__, ':', e)
                data = None
                messagebox.showerror(title="错误", message="不是十六进制！")
        return data

    """
    转换文本框中的字符为hex或str
    """

    def text_hex_convert(self, text, hex: bool):
        result = True
        op = text["state"] == DISABLED
        temp = text.get('1.0', tk.END)[:-1]
        if op:
            text.config(state=NORMAL)
        text.delete('1.0', tk.END)
        if hex:
            new_val = temp.encode("utf-8").hex()
        else:
            try:
                new_val = bytes.fromhex(temp).decode("utf-8")
            except:
                new_val = temp
                print("转换失败")
                result = False
        text.insert('1.0', new_val)
        if op:
            text.config(state=DISABLED)
        return result

    """
    十六进制复选框回调函数
    """

    def check_button_hex(self, var, cb_linebreak: tuple, text, send: bool):
        if var.get():  # 转换为十六进制
            self.text_hex_convert(text, True)
            if send:
                self.is_hex = True
                print("设置发送十六进制")
            else:
                self.to_hex = True
                """禁用接收换行"""
                cb_linebreak[1].set(False)
                self.check_button_linebreak(cb_linebreak[1], False)
                cb_linebreak[0].config(state=DISABLED)
                print("设置接收十六进制")

        else:  # 转换为str
            if self.text_hex_convert(text, False):
                if send:
                    self.is_hex = False
                    print("取消发送十六进制")
                else:
                    self.to_hex = False
                    cb_linebreak[0].config(state=NORMAL)
                    print("取消接收十六进制")
            else:
                var.set(True)  # 保持复选框状态不变
                messagebox.showerror("错误", "不是十六进制！")

    """
    定时发送复选框回调函数
    """

    def check_button_send_inerval(self, var, spinbox1, text_send, button_send):
        if var.get():
            button_send.config(state=DISABLED)
            spinbox1.config(state=DISABLED)
            if self.ser_on:
                t = threading.Thread(target=self.thread_interval_send_fun, args=(var, spinbox1, text_send))
                t.setDaemon(True)
                t.start()
        else:
            button_send.config(state=NORMAL)
            spinbox1.config(state=NORMAL)

    """
    定时发送线程函数
    """

    def thread_interval_send_fun(self, var, spinbox1, text_send):
        print('线程3启动')
        while self.ser_on:  # ser_on是全局变量
            if var.get():
                self.usart_sent(text_send.get("1.0", END))
                time.sleep(int(spinbox1.get()) / 1000)
            else:
                break
        # var.set(False)
        print('线程3结束')

    def ui_run(self):
        init_window = Tk()
        init_window.title('串口调试助手')
        #     init_window.geometry("800x700")

        frame_root = Frame(init_window)
        frame_right = Frame(frame_root)
        frame_left = Frame(frame_root)

        pw1 = PanedWindow(frame_right, orient=VERTICAL)
        pw2 = PanedWindow(frame_right, orient=VERTICAL)
        pw4 = PanedWindow(frame_right, orient=VERTICAL)

        frame1 = LabelFrame(pw1, text="串口设置")
        frame2 = Frame(frame_right)
        frame3 = LabelFrame(pw2, text="收发设置")
        frame8 = LabelFrame(pw4, text="定时发送")

        pw1.add(frame1)
        pw1.pack(side=TOP)
        frame2.pack(side=TOP)
        pw2.add(frame3)
        pw2.pack(side=TOP)
        pw4.add(frame8)
        pw4.pack()
        frame5 = Frame(frame_left)
        frame5.pack(side=TOP)
        frame6 = Frame(frame_left)
        frame6.pack(side=LEFT)
        frame7 = Frame(frame_left)
        frame7.pack(side=LEFT)

        """设置接收文本框和发送文本框"""
        text_recv = Text(frame5, width=100, height=32)
        """设置滚动条"""
        scrollbar_recv = Scrollbar(frame5)
        scrollbar_recv.pack(side=RIGHT, fill=tk.Y)
        scrollbar_recv.config(command=text_recv.yview)
        text_recv.config(yscrollcommand=scrollbar_recv.set)
        text_recv.config(state=DISABLED)
        text_recv.pack(side=LEFT)

        text_send = Text(frame6, width=84, height=15)
        text_send.grid(column=0, row=0)

        """清空接收和发送按钮"""
        frame_button = Frame(frame6)
        frame_button.grid(column=1, row=0)
        button_clear_recv = Button(frame_button, text="清空接收", width=14, height=5)
        button_clear_recv.bind('<ButtonRelease-1>', lambda event: self.text_recv_clear(text_recv))
        button_clear_recv.grid(column=0, row=0)

        button_send = Button(frame_button, text="发送", width=14, height=5)
        button_send.config(state=DISABLED)
        button_send.bind("<ButtonRelease-1>", lambda event: self.usart_sent(var=text_send.get("1.0", END)))
        button_send.grid(column=0, row=1)

        label1 = Label(frame1, text="串口号", height=2)
        label1.grid(column=0, row=0)
        label2 = Label(frame1, text="波特率", height=2)
        label2.grid(column=0, row=1)
        label3 = Label(frame1, text="数据位", height=2)
        label3.grid(column=0, row=2)
        label4 = Label(frame1, text="校验位", height=2)
        label4.grid(column=0, row=3)
        label5 = Label(frame1, text="停止位", height=2)
        label5.grid(column=0, row=4)

        """
        定时发送设置
        """
        var_sp = IntVar()
        var_sp.set(1000)
        spinbox1 = Spinbox(frame8, from_=0, to=999999, increment=1, textvariable=var_sp, width=10, justify=RIGHT)
        spinbox1.grid(column=0, row=1)

        var_checkbutton10 = BooleanVar()
        checkbutton10 = Checkbutton(frame8, variable=var_checkbutton10,
                                    command=lambda: self.check_button_send_inerval
                                    (var_checkbutton10, spinbox1, text_send, button_send))
        checkbutton10.grid(column=0, row=0)
        checkbutton10.config(state=DISABLED)
        label20 = Label(frame8, text="定时发送", width=10, height=1)
        label20.grid(column=1, row=0)

        label21 = Label(frame8, text="ms", width=8, height=1)
        label21.grid(column=1, row=1)

        """
        串口开关按钮
        """
        var_on_off = StringVar()
        var_on_off.set("打开串口")
        button_on_off = Button(frame2, textvariable=var_on_off, width=20, height=4)
        button_on_off.bind("<ButtonRelease-1>", lambda event:
        self.usart_ctrl(var_on_off, text_recv, (button_send, checkbutton10), (
            combo_port, combo_baudrate, combo_bytesize, combo_parity, combo_stopbits)))
        button_on_off.grid(column=0, row=0, pady=2)

        """
        保存按钮
        """
        button_save = Button(frame2, text="保存到文件", width=20, height=2)
        button_save.bind("<ButtonRelease-1>", lambda event: self.button_save_fun(text_recv))
        button_save.grid(column=0, row=1, pady=2)

        """
        从文件加载按钮
        """
        button_save = Button(frame2, text="从文件载入", width=20, height=2)
        button_save.bind("<ButtonRelease-1>", lambda event: self.button_load_fun(text_send))
        button_save.grid(column=0, row=2, pady=2)

        """串口下拉框"""
        varPort = StringVar()
        combo_port = ttk.Combobox(frame1, textvariable=varPort, width=12, height=10, justify=CENTER)
        """串口下拉框的串口刷新线程"""
        thread_comport_update = threading.Thread(target=self.thread_comport_update_fun, args=(combo_port,))
        thread_comport_update.setDaemon(True)
        thread_comport_update.start()
        # 以下两种写法等价?
        # combo_port['values'] = serial_com
        # combo_port.config(values=self.serial_com)
        # print(dict(combo_port))
        #     print(__file__,sys._getframe().f_lineno,m,serial_com)

        combo_port.bind("<<ComboboxSelected>>", lambda event: self.combo_port_handler(varPort.get(), var_on_off))
        # combo_port.current(0)
        combo_port.grid(column=1, row=0)

        """波特率下拉框"""
        varBaudrate = StringVar()
        combo_baudrate = ttk.Combobox(frame1, textvariable=varBaudrate, width=12, height=10, justify=CENTER)
        combo_baudrate['values'] = (
            "50", "75", "110", "134", "150", "200", "300", "600", "1200", "1800", "2400", "4800", "9600",
            "19200", "38400", "57600", "115200", "230400", "460800", "500000", "576000", "921600",
            "1000000", "1152000", "1500000", "2000000", "2500000", "3000000", "3500000", "4000000")
        combo_baudrate.bind("<<ComboboxSelected>>", lambda event: self.combo_baudrate_handler(varBaudrate))
        varBaudrate.set("9600")
        combo_baudrate.grid(column=1, row=1)

        """数据位下拉框"""
        varBytesize = StringVar()
        combo_bytesize = ttk.Combobox(frame1, textvariable=varBytesize, width=12, height=4, justify=CENTER)
        combo_bytesize['values'] = ("5", "6", "7", "8")
        combo_bytesize.bind("<<ComboboxSelected>>", lambda event: self.combo_bytesize_handler(varBytesize))
        varBytesize.set("8")
        combo_bytesize.grid(column=1, row=2)

        """校验位下拉框"""
        varParity = StringVar()
        combo_parity = ttk.Combobox(frame1, textvariable=varParity, width=12, height=5, justify=CENTER)
        combo_parity['values'] = ("NONE", "ODD", "EVEN", "MARK", "SPACE")
        combo_parity.bind("<<ComboboxSelected>>", lambda event: self.combo_parity_handler(varParity))
        varParity.set("NONE")
        combo_parity.grid(column=1, row=3)

        """停止位下拉框"""
        varStopbits = StringVar()
        combo_stopbits = ttk.Combobox(frame1, textvariable=varStopbits, width=12, height=3, justify=CENTER)
        combo_stopbits['values'] = ("1", "1.5", "2")
        combo_stopbits.bind("<<ComboboxSelected>>", lambda event: self.combo_stopbits_handler(varStopbits))
        varStopbits.set("1")
        combo_stopbits.grid(column=1, row=4)

        """
        接收发送设置
        """
        val_checkbutton1 = BooleanVar()
        val_checkbutton2 = BooleanVar()
        val_checkbutton3 = BooleanVar()
        val_checkbutton4 = BooleanVar()

        checkbutton1 = Checkbutton(frame3, variable=val_checkbutton1,
                                   command=lambda: self.check_button_linebreak(val_checkbutton1, False))
        checkbutton2 = Checkbutton(frame3, variable=val_checkbutton2,
                                   command=lambda: self.check_button_linebreak(val_checkbutton2, True))
        checkbutton3 = Checkbutton(frame3, variable=val_checkbutton3,
                                   command=lambda: self.check_button_hex(val_checkbutton3,
                                                                         (checkbutton1, val_checkbutton1), text_recv,
                                                                         False))
        checkbutton4 = Checkbutton(frame3, variable=val_checkbutton4,
                                   command=lambda: self.check_button_hex(val_checkbutton4, (), text_send, True))
        label7 = Label(frame3, text="接收时添加新行", width=16, height=1, justify=LEFT)
        label8 = Label(frame3, text="发送时添加新行", width=16, height=1, justify=LEFT)
        label9 = Label(frame3, text="十六进制接收", width=16, height=1, justify=LEFT)
        label10 = Label(frame3, text="十六进制发送", width=16, height=1, justify=LEFT)

        checkbutton1.grid(column=0, row=0)
        checkbutton2.grid(column=0, row=1)
        checkbutton3.grid(column=0, row=2)
        checkbutton4.grid(column=0, row=3)

        label7.grid(column=1, row=0)
        label8.grid(column=1, row=1)
        label9.grid(column=1, row=2)
        label10.grid(column=1, row=3)

        """
        窗口框架
        """

        frame_right.pack(side=RIGHT)
        frame_left.pack(side=LEFT)
        frame_root.pack()

        """关闭窗口动作"""
        init_window.protocol('WM_DELETE_WINDOW', lambda: self.closing_procedure(init_window.destroy))
        """ui循环"""
        init_window.mainloop()


if __name__ == "__main__":
    ui = ComUI()
    ui.ui_run()
