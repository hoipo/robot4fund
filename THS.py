from pywinauto.application import Application
from pywinauto.timings import wait_until, TimeoutError
from pywinauto import keyboard, clipboard
from time import sleep
import io
import pandas as pd

class THS:
    def __init__(self, exe_path=r"C:\同花顺软件\同花顺\xiadan.exe"):
        """
        init the class of THS
        """
        print("正在连接客户端:", exe_path, "......")
        self.app = Application().connect(path=exe_path, timeout=10)
        self.main_win = self.app.window(title="网上股票交易系统5.0")
        print("连接成功!!!")

    def sell(self, symbol, money, limited_price, all_accounts=False):
        """
        卖出股票或基金
        """
        self.__select_menu(path="\卖出[F2]")
        sleep(1)
        result = self.__trade(symbol=symbol, limited_price=limited_price, money=money)
        print(result)

    def clear_position(self, symbol, limited_price):
        """
        清仓
        """
        self.__select_menu(path="\卖出[F2]")
        self.__esc_extra_dialog()
        for i in range(self.main_win.window(control_id=0x3ec).item_count()):
            self.__select_account(i)
            result = self.__trade(symbol=symbol, limited_price=limited_price)
            print(result)

    def __esc_extra_dialog(self):
        """
        退出意外的对话框
        """
        while True:
            sleep(1)
            if self.main_win.get_active().window_text() != "网上股票交易系统5.0":
                keyboard.send_keys("{ESC}")
            else:
                break

    def __trade(self, symbol, limited_price, money=None):
        self.__esc_extra_dialog()
        sleep(0.5)
        self.main_win.window(control_id=0x408, class_name="Edit").set_focus()  # 设置股票代码
        keyboard.send_keys(symbol)
        sleep(0.5)
        self.__select_stock_market()
        sleep(0.5)
        self.main_win.window(control_id=0x409, class_name="Edit").set_focus()  # 设置价格
        keyboard.send_keys("{END}" "{BACKSPACE 20}")
        keyboard.send_keys(limited_price)

        if money is None:
            money = self.main_win.window(control_id=0x40E, class_name="Static").window_text()
        if int(money) > 0:
            self.main_win.window(control_id=0x40A, class_name="Edit").set_focus()  # 设置股数目
            sleep(0.5)
            self.__select_stock_market()
            sleep(0.5)
            self.main_win.window(control_id=0x40A, class_name="Edit").set_focus()  # 设置股数目
            keyboard.send_keys(money)
            sleep(0.5)
            self.main_win.window(control_id=0x3EE, class_name="Button").click()   # 点击卖出or买入

            sleep(0.5)
            keyboard.send_keys('{SPACE}')

            sleep(1)
            result = self.main_win.get_active().children(control_id=0x3EC, class_name='Static')[0].window_text()

            sleep(0.5)
            try:
                self.main_win.get_active().children(control_id=0x2, class_name='Button')[0].set_focus()  # 确定
                keyboard.send_keys("{SPACE}")
            except:
                pass
            return self.__parse_result(result)
        else:
            return {'success': False, 'msg': ''}

    @staticmethod
    def __parse_result(result):
        """ 解析买入卖出的结果 """
        if r"已成功提交，合同编号：" in result:
            return {
                "success": True,
                "msg": result,
                "entrust_no": result.split("合同编号：")[1].split("。")[0]
            }
        else:
            return {
                "success": False,
                "msg": result
            }

    def __select_stock_market(self):
        """
        选择证券市场
        """
        # 选择证券，如果需要的话
        temp_win = self.main_win.get_active()
        if len(temp_win.children(title="请选择证券")) > 0:
            for child in temp_win.iter_children():
                if 'Ａ股' in child.window_text():
                    child.click()
                    break


    def purchase_fund(self, fund_code, money):
        """
        申购场内基金
        """
        self.__select_menu(path="\场内基金\基金申购")
        for i in range(4):
            keyboard.send_keys("{TAB}")

        for i in range(self.main_win.window(control_id=0x3ec).item_count()):
            keyboard.send_keys(fund_code)
            sleep(0.5)
            self.__select_stock_market()
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

    def __get_grid_data(self, is_records=True):
        """ 获取grid里面的数据 """
        self.__click_update_button()
        
        sleep(0.5)
        self.__esc_extra_dialog()
        grid = self.main_win.window(control_id=0x417, class_name='CVirtualGridCtrl')
        grid.set_focus()
        sleep(0.5)
        keyboard.send_keys('^c')
        data = clipboard.GetData()
        df = pd.read_csv(io.StringIO(data), delimiter='\t', na_filter=False)
        if is_records:
            return df.to_dict('records')
        else:
            return df

    def __click_update_button(self):
        print(self.main_win.window(control_id=0xE800, class_name='ToolbarWindow32').press_button(3))

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
        self.__get_left_menus_handle().print_items()
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
