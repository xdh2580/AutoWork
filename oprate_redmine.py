from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select



redmine_adress = "192.168.3.78:8078"
username = ""
password = ""
project = "xqt551_a13-go"

DEFAULT_SOFTWARE_VERSION = r"\\192.168.4.3\Software\MK2\02_FirstTA\01_First_version"
DEFAULT_NOTES = "单测报告及log在附件"
DEFAULT_ASSIGN_TO = "huangzhangbin黄章斌"
DEFAULT_ANDROID_VERSION = "Android13"

def init(driver):
    driver.maximize_window()
    driver.get("http://" + redmine_adress + "/login")
    driver.find_element(By.ID, value="username").send_keys(username)
    driver.find_element(By.ID, value="password").send_keys(password)
    driver.find_element(By.ID, value="login-submit").click()
    # driver.get("http://192.168.3.78:8078/login")



def fill_content(driver,xTS:str, module:str, num:str, case:str, tool:str):
    isuue_subject = driver.find_element(By.ID, value="issue_subject")  # 标题
    isuue_subject.send_keys(f"[BUG]({xTS}){module}模块存在{num}条失败项")
    issue_description = driver.find_element(By.ID, value="issue_description")  # 描述
    issue_description.send_keys(f"【模块】：{module}\n【case】：{case}\n【测试工具】：{tool}\n【软件版本】：{DEFAULT_SOFTWARE_VERSION}\n【备注】：{DEFAULT_NOTES}")
    issue_assigned_to_id = Select(driver.find_element(By.ID, value="issue_assigned_to_id"))  # 指派给
    issue_assigned_to_id.select_by_visible_text(DEFAULT_ASSIGN_TO)
    issue_custom_field_values_1 = Select(driver.find_element(By.ID, value="issue_custom_field_values_1"))  # 软件平台
    issue_custom_field_values_1.select_by_visible_text(DEFAULT_ANDROID_VERSION)
    issue_custom_field_values_3 = Select(driver.find_element(By.ID, value="issue_custom_field_values_3"))  # 问题涉及模块
    issue_custom_field_values_3.select_by_visible_text("GMS")
    issue_custom_field_values_35 = Select(driver.find_element(By.ID, value="issue_custom_field_values_35"))  # bug版本
    issue_custom_field_values_35.select_by_visible_text("V00")
    issue_custom_field_values_4 = Select(driver.find_element(By.ID, value="issue_custom_field_values_4"))  # 问题或任务类别
    issue_custom_field_values_4.select_by_visible_text("问题反馈Issue")
    issue_custom_field_values_5 = Select(driver.find_element(By.ID, value="issue_custom_field_values_5"))  # bug类别
    issue_custom_field_values_5.select_by_visible_text("BUG")
    issue_custom_field_values_7 = Select(driver.find_element(By.ID, value="issue_custom_field_values_7"))  # bug缺陷等级
    issue_custom_field_values_7.select_by_visible_text("20--主要缺陷（Major）")
    issue_custom_field_values_12 = Select(driver.find_element(By.ID, value="issue_custom_field_values_12"))  # bug可见性
    issue_custom_field_values_12.select_by_visible_text("5--一般见(Normal)")
    issue_custom_field_values_16 = Select(driver.find_element(By.ID, value="issue_custom_field_values_16"))  # bug复现概率
    issue_custom_field_values_16.select_by_visible_text("10--必现(Always)")
    issue_custom_field_values_17 = Select(driver.find_element(By.ID, value="issue_custom_field_values_17"))  # bug等级
    issue_custom_field_values_17.select_by_visible_text("B (1000>= FMEA得分 >600)")
    issue_custom_field_values_18 = driver.find_element(By.ID, value="issue_custom_field_values_18")  # bug_FMEA得分
    issue_custom_field_values_18.send_keys("1000")

#all_info:所有报告信息字典组成的列表
def new_all_bugs(all_info, driver):
    init(driver)
    current_window = 0
    for info in all_info:
        plan = info["suite_plan"].split(" / ")[0]
        if info["suite_plan"].split(" / ")[1] == "cts-on-gsi":
            plan = "CTS-ON-GSI"
        build = info["suite_build"]
        tool = plan+build
        if plan == "CTS-ON-GSI":
            tool = "CTS"+build
        if plan == "STS":
            tool = build
        print("tool:"+tool)
        fails = info["fails"]  # info["fails"]是个列表,其中每个fail元素是字典
        fails_order_by_module = {}  # 按模块归类的列表--每个元素是字典，其module为模块名，v7a及v8a和[instant]归到同一moudule,case为该模块下fail case的集合
        # fails_in_module = {"module": module, "case": set()}
        for fail in fails:
            module = fail["module"].replace("[instant]", "")
            module = module.split(" ")[1]
            if fails_order_by_module.get(module, 1) == 1:
                fails_order_by_module[module] = set()
            fails_order_by_module[module].add(fail['name'])
        for key in fails_order_by_module:
            num = len(fails_order_by_module[key])
            # print("模块"+key+"下存在"+str(num)+"条失败项")

            #每个模块建一个bug
            new_url = "http://" + redmine_adress + "/projects/" + project + "/issues/new"
            driver.execute_script(f'window.open("{new_url}")')
            window_list = driver.window_handles
            driver.switch_to.window(window_list[current_window + 1])
            current_window += 1
            cases = ""
            for case in fails_order_by_module[key]:
                cases = cases+"\n"+case
            fill_content(driver, plan, key, str(num), cases, tool)
    input("按 Enter 键结束程序...")
