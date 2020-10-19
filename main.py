from THS import THS

exe_path = input(r"请输入同花顺下跌程序的完整目录(如：C:\同花顺软件\同花顺\xiadan.exe):") 
fund_code = input("请输入基金代码：")
money = input("请输入申购金额：")
ths = THS(exe_path=exe_path or r"C:\同花顺软件\同花顺\xiadan.exe")
ths.purchase_fund(fund_code=fund_code, money=money)

input(r'按回车键退出...')