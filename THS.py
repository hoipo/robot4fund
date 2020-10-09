from pywinauto.application import Application
from pywinauto.timings import wait_until, TimeoutError
from pywinauto import keyboard
from time import sleep

class THS:
    def __init__(self, exe_path=r"C:\同花顺软件\同花顺\xiadan.exe"):
        """
        init the class of THS
        """
        print("正在连接客户端:", exe_path, "......")
        self.app = Application().connect(path=exe_path, timeout=10)
        self.main_win = self.app.window(title="网上股票交易系统5.0")
        print("连接成功!!!")

    def purchase_fund(self, fund_code, money):
        """
        申购场内基金
        """
        self.__select_menu(path="\场内基金\基金申购")
        for i in range(4):
            keyboard.send_keys("{TAB}")

        for i in range(self.main_win.window(control_id=0x3ec).item_count()):
            keyboard.send_keys(fund_code)
            self.__select_account(index=i)
            keyboard.send_keys("{TAB}")
            keyboard.send_keys(money)
            keyboard.send_keys("%s" "%y" "{UP}" "{SPACE}")
            
            wait_until(10, .5, self.__get_target_dialog, True,
                    text_in_dialog="本人已经认真阅读并理解上述内容")  # 等待对话框
            keyboard.send_keys("{TAB}" "{SPACE}" "{TAB}" "{ENTER}")

            wait_until(10, .5, self.__get_target_dialog, True,
                    text_in_dialog="提示信息")  # 等待对话框
            keyboard.send_keys("{ENTER}")
            
            wait_until(10, .5, self.__get_target_dialog, True,
                    text_in_dialog="公募证券投资基金投资风险告知")  # 等待对话框
            keyboard.send_keys("{TAB}" "{SPACE}" "{TAB}" "{ENTER}")
            
            wait_until(10, .5, self.__get_target_dialog, True,
                    text_in_dialog="适当性匹配结果确认书")  # 等待对话框
            keyboard.send_keys("{TAB}" "{ENTER}")
            
            wait_until(10, .5, self.__get_target_dialog, True,
                    text_in_dialog="提示")  # 等待对话框
            keyboard.send_keys("{ENTER}")

    def __select_account(self, index):
        """
        选择股东账户, index是股东账户的序号
        """
        current_focus_element = self.main_win.get_focus()
        self.main_win.window(control_id=0x3ec).set_focus()
        keyboard.send_keys("{HOME}")
        for j in range(index):
            keyboard.send_keys("{RIGHT}")
        current_focus_element.set_focus()

    def __get_target_dialog(self, text_in_dialog):
        """
        判断对话框是否存在
        """
        dgl = self.main_win.get_active()
        return len(dgl.children(title=text_in_dialog)) > 0

    def __select_menu(self, path):
        """ 点击左边菜单 """
        self.main_win.set_focus()
        keyboard.send_keys("{F1}")
        self.__get_left_menus_handle().select(path=path)

    def __get_left_menus_handle(self):
        while True:
            try:
                handle = self.main_win.window(class_name="SysTreeView32")
                handle.wait('ready', timeout=2)
                return handle
            except Exception as ex:
                print(ex)
                pass
