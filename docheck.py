from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import os
from openpyxl import load_workbook
import _thread
import tkinter
from tkinter.messagebox import showinfo


# path：报告路径，必须是下一级包含"cts","vts"等文件夹的目录，也即一版软件的报告路径
# 返回所有test_resule_failure报告的路径的列表
def getreport(path):
    list_dir = os.listdir(path)
    path_report = []
    for i in list_dir:
        if i == "cts" or i == "vts" or i == "sts" or i == "gts" or i == "gsi" or i == "cts-instant":
            sub_list = os.path.join(path, i)
            try:
                for j in os.listdir(sub_list):
                    if j[:3] == "202":
                        sub_list2 = os.path.join(sub_list, j)
                        try:
                            for q in os.listdir(sub_list2):
                                if q == "test_result_failures_suite.html" or q == "test_result_failures.html":
                                    path_report.append(os.path.join(sub_list2, q))
                        except Exception as e:
                            print("非目录，跳过")
            except Exception as e:
                print("跳过")
    return path_report


# report:报告文件的路径
# 返回该报告中的一些信息
def getinfo(report):
    infodict = dict()
    # 设置selenium使用chrome的无头模式
    chrome_options = Options()

    chrome_options.add_argument('headless')  # 设置option,隐藏浏览器界面
    # 在启动浏览器时加入配置
    browser = webdriver.Chrome(r'chromedriver.exe', options=chrome_options)  # 获取chrome浏览器的驱动，并启动Chrome浏览器
    browser.get('file:///' + report)
    # 等待加载，最多等待20秒
    browser.implicitly_wait(20)
    # browser.maximize_window()  # 窗口最大化

    li = browser.find_elements_by_xpath("//td[@class='rowtitle']/../td[2]")
    infodict["suite_plan"] = li[0].text
    infodict["suite_build"] = li[1].text
    infodict["host_info"] = li[2].text
    infodict["time_info"] = li[3].text
    infodict["case_pass"] = li[4].text
    infodict["case_fail"] = li[5].text
    infodict["modules_done"] = li[6].text
    infodict["modules_total"] = li[7].text
    infodict["finger_print"] = li[8].text
    infodict["security_patch"] = li[9].text
    infodict["release_sdk"] = li[10].text
    infodict["ABIs"] = li[11].text

    details = browser.find_elements_by_class_name("testdetails")
    all_fail = []  # 所有fail项的列表
    for fail_module in details:
        module_name = fail_module.find_element_by_class_name("module").text
        fails = fail_module.find_elements_by_class_name("testname")
        for fail_ietm in fails:
            fail = {}  # 每个fail为字典，module：所属模块，name：失败项case名称，detail：报错信息
            fail["module"] = module_name
            fail["name"] = fail_ietm.text
            fail["detail"] = "null for temp"  # 待补充
            all_fail.append(fail)
    infodict["fails"] = all_fail
    # print(all_fail)
    browser.quit()
    return infodict


# 在表格模板中填充信息
# all_info：所有报告信息字典的列表
# path：汇总表格文件的输出目录
def write_xl(all_info, path=""):
    workbook = load_workbook(filename="case.xlsx")
    sheet1 = workbook["case汇总"]
    sheet2 = workbook["失败项汇总"]
    row0 = ["plan", "tool", "case_all_test", "case_pass", "case_fail", "modules_done", "modules_total", "finger_print"]
    sheet1.append(row0)
    data = []
    for info in all_info:
        plan = info["suite_plan"]
        build = info["suite_build"]
        case_pass = info["case_pass"]
        case_fail = info["case_fail"]
        modules_done = info["modules_done"]
        modules_total = info["modules_total"]
        finger_print = info["finger_print"]

        fails = info["fails"]#info["fails"]是个列表,其中每个fail元素是字典
        for fail in fails:
            row_fail = [plan, fail["module"], fail["name"]]  # , fail["detail"]
            sheet2.append(row_fail)

        case_all_test = int(case_pass) + int(case_fail)
        row = [plan, build, case_all_test, int(case_pass), int(case_fail), int(modules_done), int(modules_total),
               finger_print]
        data.append(row)  # 将信息直接附在后面
        #  2021.10.19 start 自动填充工具信息及模块和case数到模板中的固定位置，同一plan多个报告取total_case数量最多的
        p = plan.split('/')
        s = build.split('/')
        tool = p[0] + s[0]
        if plan == "CTS / cts" or plan == "CTS / cts-retry":
            sheet1['B1'] = tool
            sheet1['C1'] = tool
            if sheet1["C3"].value is None or int(modules_total) > sheet1["C3"].value:
                sheet1["C3"] = int(modules_total)
            if sheet1["C4"].value is None or int(modules_total) > sheet1["C4"].value:
                sheet1["C4"] = case_all_test
        if plan == "VTS / cts-on-gsi" or plan == "VTS / cts-on-gsi-retry":
            sheet1['D1'] = tool
            if sheet1["D3"].value is None or int(modules_total) > sheet1["D3"].value:
                sheet1["D3"] = int(modules_total)
            if sheet1["D4"].value is None or int(modules_total) > sheet1["D4"].value:
                sheet1["D4"] = case_all_test
        if plan == "VTS / vts":
            sheet1['E1'] = tool
            if sheet1["E3"].value is None or int(modules_total) > sheet1["E3"].value:
                sheet1["E3"] = int(modules_total)
            if sheet1["E4"].value is None or int(modules_total) > sheet1["E4"].value:
                sheet1["E4"] = case_all_test
        if plan == "GTS / gts":
            sheet1['F1'] = tool
            if sheet1["F3"].value is None or int(modules_total) > sheet1["F3"].value:
                sheet1["F3"] = int(modules_total)
            if sheet1["F4"].value is None or int(modules_total) > sheet1["F4"].value:
                sheet1["F4"] = case_all_test
        if plan == "STS / sts-engbuild":
            sheet1['G1'] = tool
            if sheet1["G3"].value is None or int(modules_total) > sheet1["G3"].value:
                sheet1["G3"] = int(modules_total)
            if sheet1["G4"].value is None or int(modules_total) > sheet1["G4"].value:
                sheet1["G4"] = case_all_test
        #  2021.10.19 end
    for r in data:
        sheet1.append(r)
    workbook.save(filename=path + r"\汇总.xlsx")


def do_my_print(path):
    _thread.start_new_thread(real_do, (path,))


def real_do(path):
    label2 = tkinter.Label(main_window, text="执行中...")
    label2.pack()
    print("path: " + path)
    list_dir = os.listdir(path)
    sub = False
    for i in list_dir:
        if i == "CN" or i == "EU" or i == "RU" or i == "US":
            sub = True
            path_sub = os.path.join(path, i)
            print("path_sub:" + path_sub)
            real_real_do(path_sub)
    if not sub:
        real_real_do(path)
    showinfo(title="完成", message="完成！汇总表格已保存：\n" + path)
    label2.pack_forget()


def real_real_do(path):
    all_info = []  # 所有报告的信息字典的列表
    for i in getreport(path):
        infodict = getinfo(i)
        all_info.append(infodict)
        print("dict:" + str(infodict))
    write_xl(all_info, path)
    print("完成！" + path)


def init_window():
    main_window.title("Auto Check")
    main_window.geometry("300x200")
    label1 = tkinter.Label(main_window, text="请输入路径：")
    label1.pack()
    entry1 = tkinter.Entry(main_window)
    entry1.pack()
    button1 = tkinter.Button(main_window, text="开始", command=lambda: do_my_print(entry1.get()))
    button1.pack()


if __name__ == '__main__':
    main_window = tkinter.Tk()
    init_window()
    main_window.mainloop()
