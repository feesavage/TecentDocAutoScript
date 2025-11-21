# -*- coding: utf-8 -*-
"""
腾讯文档自动填写脚本

功能说明：
- 使用Selenium自动化工具自动填写腾讯文档表单
- 支持定时执行，可在指定时间自动开始填写
- 自动检测并填写表单中的文本输入框
- 支持下拉菜单自动选择功能
- 自动处理表单提交和确认操作

使用方法：
1. 修改 url 变量为目标腾讯文档地址
2. 修改 execute_time 变量为执行时间
3. 修改 elements[i].send_keys() 中的内容为实际要填写的值
4. 如需选择下拉菜单，使用以下方法：
   - 通过文本选择：select_dropdown_option("选项名称")
   - 通过索引选择：select_dropdown_by_index(索引号)
5. 运行脚本，手动扫码登录后输入 'y' 开始等待执行

下拉菜单使用示例：
- select_dropdown_option("男")  # 选择文本为"男"的选项
- select_dropdown_by_index(1)   # 选择第2个选项（索引从0开始）

注意事项：
- 需要安装 Chrome 浏览器和对应版本的 ChromeDriver
- 确保网络连接稳定
- 请提前测试表单结构，确保选择器正确
- 下拉菜单选择器可能需要根据具体页面结构调整
"""

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import datetime


# ====================== 下拉菜单选择功能定义 ======================
def select_dropdown_option(dropdown_selector, option_text):
    """
    选择指定下拉菜单中包含特定文本的选项

    参数:
        dropdown_selector (str): 下拉菜单的选择器
        option_text (str): 要选择的选项文本

    返回:
        bool: 是否成功选择选项
    """
    try:
        # 等待并点击指定的下拉菜单触发器
        dropdown_trigger = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, dropdown_selector))
        )
        dropdown_trigger.click()
        time.sleep(0.5)  # 等待下拉菜单展开

        # 尝试多种方式查找选项
        option_selectors = [
            f"//div[contains(@class, 'dropdown-option') and contains(text(), '{option_text}')]",
            f"//li[contains(@class, 'dropdown-option') and contains(text(), '{option_text}')]",
            f"//div[contains(@data-dui, 'dropdown-option') and contains(text(), '{option_text}')]",
            f"//li[contains(@data-dui, 'dropdown-option') and contains(text(), '{option_text}')]",
            f"//*[contains(@class, 'option') and contains(text(), '{option_text}')]",
            f"//div[contains(text(), '{option_text}') and not(contains(text(), 'placeholder'))]",
        ]

        for selector in option_selectors:
            try:
                option = WebDriverWait(driver, 2).until(
                    EC.element_to_be_clickable((By.XPATH, selector))
                )
                option.click()
                print(f"已选择下拉菜单选项：{option_text}")
                return True
            except:
                continue

        print(f"未找到包含文本 '{option_text}' 的选项")
        return False

    except Exception as e:
        print(f"选择下拉菜单选项 '{option_text}' 失败: {e}")
        return False

def select_dropdown_by_index(dropdown_selector, index):
    """
    通过索引选择指定下拉菜单的选项

    参数:
        dropdown_selector (str): 下拉菜单的选择器
        index (int): 选项的索引（从0开始）

    返回:
        bool: 是否成功选择选项
    """
    try:
        # 等待并点击指定的下拉菜单触发器
        dropdown_trigger = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, dropdown_selector))
        )
        dropdown_trigger.click()
        time.sleep(0.5)  # 等待下拉菜单展开

        # 尝试多种方式查找所有选项
        option_selectors = [
            "div[class*='dropdown-option']",
            "li[class*='dropdown-option']",
            "div[data-dui*='dropdown-option']",
            "li[data-dui*='dropdown-option']",
            "[class*='option']",
        ]

        options = []
        for selector in option_selectors:
            options = driver.find_elements(By.CSS_SELECTOR, selector)
            if options:
                break

        if index < len(options):
            options[index].click()
            print(f"已选择下拉菜单第 {index + 1} 个选项")
            return True
        else:
            print(f"索引 {index} 超出选项范围，共有 {len(options)} 个选项")
            return False

    except Exception as e:
        print(f"通过索引 {index} 选择下拉菜单选项失败: {e}")
        return False

def find_and_select_all_dropdowns():
    """
    自动查找页面上的所有下拉菜单

    返回:
        list: 找到的下拉菜单元素列表
    """
    dropdown_selectors = [
        "[data-dui*='dropdown']",
        "[class*='dropdown']",
        "[class*='select']",
        "select",
    ]

    dropdowns = []
    for selector in dropdown_selectors:
        elements = driver.find_elements(By.CSS_SELECTOR, selector)
        if elements:
            dropdowns.extend(elements)

    print(f"找到 {len(dropdowns)} 个下拉菜单")
    return dropdowns


# 初始化浏览器配置
options = webdriver.ChromeOptions()
# 如果需要无头模式，可以取消下面这行注释
# options.add_argument('--headless')

# 启动Chrome浏览器
driver = webdriver.Chrome(options=options)

# 目标腾讯文档URL（需要用户修改）
url = "https://docs.qq.com/form/page/DZWR3aUVLaWh2Uk9U#/fill"
driver.get(url)

# 等待用户手动登录确认
# 由于腾讯文档需要扫码登录，这里暂停等待用户确认登录完成
while True:
    ch = input("是否已登录完成？(y/n): ")
    if ch == 'y':
        break

# 注释掉的代码：自动登录功能（可选，根据实际情况启用）
# elmet = driver.find_element(By.ID,"header-login-btn")
# elmet.click()
# driver.implicitly_wait(2)
# elmet=driver.find_element(By.CSS_SELECTOR,'span.qq')
# elmet.click()

# 设定执行时间（需要用户修改为实际执行时间）
# 格式：datetime.datetime(年, 月, 日, 小时, 分钟, 秒)
execute_time = datetime.datetime(2024, 10, 23, 22, 50, 0)

# 计算距离执行时间的等待时间
wait_time = execute_time - datetime.datetime.now()

# 等待到指定时间再执行
if wait_time.total_seconds() > 0:
    print("等待到指定执行时间...")
    time.sleep(wait_time.total_seconds())

print("开始执行！")
start_time = time.time()

# 刷新页面确保获取最新状态
driver.execute_script("window.location.reload()")



# 查找并填写表单中的文本输入框
# 注意：XPath选择器 "//textarea[]" 可能需要根据实际页面结构调整
timeout = 10  # 元素查找超时时间（秒）
locator = (By.XPATH, "//textarea[]")
elements = []

# 等待直到找到文本输入框元素
while not elements:
    try:
        elements = WebDriverWait(driver, timeout).until(
            EC.presence_of_all_elements_located(locator)
        )
    except:
        print("未找到输入框，继续等待...")
        # 再次刷新
        driver.execute_script("window.location.reload()")
        time.sleep(0.5)

# 填写表单内容（需要用户修改为实际内容）
# elements[0] 通常对应第一个文本输入框
# elements[0].send_keys("11")   # 第一个输入框内容
# elements[1].send_keys("22")   # 第二个输入框内容
# 如果有更多输入框，可以取消下面的注释
# elements[2].send_keys("33")   # 第三个输入框内容
# elements[3].send_keys("44")   # 第四个输入框内容

# ====================== 选择下拉菜单选项 ======================
# 在填写完文本框后，自动查找并选择下拉菜单选项

# 自动查找页面上的下拉菜单
dropdowns = find_and_select_all_dropdowns()

# 如果找到下拉菜单，自动选择选项（根据实际需求修改）
if dropdowns:
    print(f"检测到 {len(dropdowns)} 个下拉菜单，开始自动选择...")

    # 示例1：选择第一个下拉菜单的第2个选项，index为0是第一个选项
    if len(dropdowns) >= 1:
        select_dropdown_by_index("[data-dui*='dropdown']", 1)

    # 示例2：如果需要根据文本选择，取消下面注释并修改选项文本
    # select_dropdown_option("[data-dui*='dropdown']", "男")
    # select_dropdown_option("[data-dui*='dropdown']", "是")
    # select_dropdown_option("[data-dui*='dropdown']", "同意")

    # 示例3：处理多个下拉菜单
    # for i, dropdown in enumerate(dropdowns):
    #     print(f"下拉菜单 {i+1}: 准备选择默认选项")
    #     select_dropdown_by_index(f"[data-dui*='dropdown']:nth-of-type({i+1})", 0)
else:
    print("未检测到下拉菜单，跳过选择步骤")

# 下拉菜单选择已完成，继续执行提交流程


# 第一步：点击"提交"按钮
# 使用更灵活的XPath选择器来查找包含"提交"文字的按钮
submit_button = driver.find_element(By.XPATH, "//button[text()='提交']")
driver.execute_script("arguments[0].click();", submit_button)

# 第二步：等待并点击确认按钮
# 等待确认弹窗出现，然后点击包含"确认"文字的按钮
confirm_locator = (By.XPATH, "//button[contains(.,'确认')]")
confirm_button = WebDriverWait(driver, timeout).until(
    EC.presence_of_element_located(confirm_locator)
)
confirm_button.click()

# ====================== 执行结果统计 ======================
end_time = time.time()
execution_time = end_time - start_time

print(f"表单填写完成！")
print(f"执行耗时：{execution_time:.2f} 秒")
print(f"当前时间：{datetime.datetime.now()}")

# 注意：在实际使用前，建议先手动测试XPath选择器的准确性
# 可以在浏览器开发者工具中验证选择器是否能正确找到目标元素


