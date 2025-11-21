# -*- coding: utf-8 -*-
"""
腾讯文档自动填写脚本
日期: 2025/11/21 作者：fee
功能说明：
- 使用Selenium自动化工具自动填写腾讯文档表单
- 支持定时执行，可在指定时间自动开始填写
- 自动检测并填写表单中的文本输入框
- 支持下拉菜单自动选择功能
- 支持单选按钮自动选择功能
- 自动处理表单提交和确认操作

使用方法：
1. 修改 url 变量为目标腾讯文档地址
2. 修改 execute_time 变量为执行时间
3. 修改 elements[i].send_keys() 中的内容为实际要填写的值
4. 如需选择下拉菜单，使用以下方法：
   - 通过文本选择：select_dropdown_by_text("选项名称")
   - 通过索引选择：select_dropdown_by_index(索引号)
5. 如需选择单选按钮，使用以下方法：
   - 通过文本选择：select_radio_by_text("选项名称")
   - 通过索引选择：select_radio_by_index(索引号)
6. 运行脚本，手动扫码登录后输入 'y' 开始等待执行


注意事项：
- 需要安装 Chrome 浏览器和对应版本的 ChromeDriver
- 确保网络连接稳定
- 请提前测试表单结构，确保选择器正确
- 下拉菜单和单选按钮选择器可能需要根据具体页面结构调整
"""

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, InvalidSelectorException
import time
import datetime


# 匹配信息框填写的规则
TEXT_XPATH = "//textarea"
# 匹配下拉菜单的规则
DROPDOWN_XPATH = '//*[@role="combobox"]'
# 匹配单选题的规则
RADIO_GROUP_XPATH = '//div[@class="question-content"]/div[contains(@class, "form-choice-radio")]'
# 匹配单选选项的规则
RADIO_XPATH = './/div[@role="radio"]'
# 填写完毕后，页面关闭的等待时间（s）
N = 10
#填写的URl,用户自己修改
URL = "https://docs.qq.com/form/page/DZWR3aUVLaWh2Uk9U"

def find_all_dropdowns():
    """
    自动查找页面上的所有下拉菜单
    """

    dropdowns = []
    elements = driver.find_elements(By.XPATH, DROPDOWN_XPATH)
    if elements:
        dropdowns.extend(elements)

    print(f"找到 {len(dropdowns)} 个下拉菜单")
    return dropdowns


def find_all_radios_groups():
    """
    查找页面上的所有单选题
    """
    radios_groups = []

    try:
        # 查找包含单选按钮容器

        elements = driver.find_elements(By.XPATH, RADIO_GROUP_XPATH)
        if elements:
            radios_groups.extend(elements)

        print(f"找到 {len(radios_groups)} 个单选题")

        return radios_groups

    except Exception as e:
        print(f"查找单选题失败: {e}")
        return []




def select_dropdown_by_index(dropdown,index):
    """
    通过索引选择指定下拉菜单的选项

    参数:
        dropdown：下拉菜单对象
        index (int): 选项的索引（从0开始）

    返回:
        bool: 是否成功选择选项
    """
    try:
        dropdown.click()
        time.sleep(0.5)  # 等待下拉菜单展开

        aria_owns = dropdown.get_attribute("aria-owns")
        if aria_owns:
            options = driver.find_elements(By.XPATH, f'//*[@id="{aria_owns}"]//li[@role="option"]')
        else:
            options = driver.find_elements(By.XPATH, '//li[@role="option" and contains(@class, "dui-option")]')

        enabled_options = [
            opt for opt in options
            if opt.get_attribute("aria-disabled") != "true"
        ]

        if index < len(enabled_options):
            enabled_options[index].click()
            print(f"已选择下拉菜单第 {index + 1} 个选项")
            return True
        else:
            print(f"索引 {index} 超出选项范围，共有 {len(enabled_options)} 个可用选项")
            return False

    except Exception as e:
        print(f"通过索引 {index} 选择下拉菜单选项失败: {e}")
        return False


def select_radio_by_index(radio_group,index):
    """
    在指定的单选组中选择指定索引的选项

    参数:
        radio_group: 单选对象
        index (int): 组内选项的索引（从0开始）

    返回:
        bool: 是否成功选择
    """
    try:
        options = radio_group.find_elements(By.XPATH, RADIO_XPATH)
        enabled_options = [
            opt for opt in options
            if opt.get_attribute("aria-disabled") != "true"
        ]
        if index < len(enabled_options):
            enabled_options[index].click()
            print(f"已选择第 {index + 1} 个单选项")
            return True
        else:
            print(f"索引 {index} 超出范围，共 {len(enabled_options)} 个选项")
            return False
    except Exception as e:
        print(f"通过索引 {index} 选择单选菜单选项失败: {e}")
        return False


# 初始化浏览器配置
options = webdriver.ChromeOptions()
# 如果需要无头模式，可以取消下面这行注释
# options.add_argument('--headless')

# 启动Chrome浏览器
driver = webdriver.Chrome(options=options)

driver.get(URL)

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
timeout = 10  # 元素查找超时时间（秒）
locator = (By.XPATH, TEXT_XPATH)
elements = []

# 等待直到找到文本输入框元素
while not elements:
    try:
        elements = WebDriverWait(driver, timeout).until(
            EC.presence_of_all_elements_located(locator)
        )
    except TimeoutException:
        print("未找到输入框，继续等待...")
        # 再次刷新
        driver.execute_script("window.location.reload()")
        time.sleep(0.5)
    except Exception as e:
        print(f"其他异常: {e}")
# 填写表单内容（需要用户修改为实际内容）
# elements[0] 通常对应第一个文本输入框
elements[0].send_keys("11")   # 第一个输入框内容
# elements[1].send_keys("22")   # 第二个输入框内容
# 如果有更多输入框，可以取消下面的注释
# elements[2].send_keys("33")   # 第三个输入框内容
# elements[3].send_keys("44")   # 第四个输入框内容

# ====================== 选择下拉菜单选项 ======================
# 在填写完文本框后，自动查找并选择下拉菜单选项

# 自动查找页面上的下拉菜单
dropdowns = find_all_dropdowns()

# 如果找到下拉菜单，自动选择选项（根据实际需求修改）
if dropdowns:
    print(f"检测到 {len(dropdowns)} 个下拉菜单，开始自动选择...")

    # 示例:处理多个下拉菜单选项，有几个下拉，定义几个数值，第一个数值代表问题1，index = 1 ，代表第一个选项
    # [1,0]代表总共有两个问题，第一个问题选择第二项，第二个问题选择第一项。
    indices_to_select = [0, 1]
    for i, dropdown in enumerate(dropdowns):
        print(f"下拉菜单 {i+1}: 准备选择选项")
        if i < len(indices_to_select):
            select_dropdown_by_index(dropdown,indices_to_select[i])
        else:
            print("未定义索引，选择默认第一个选项")
            select_dropdown_by_index(dropdown,0) #默认选择0
else:
    print("未检测到下拉菜单，跳过选择步骤")

# ====================== 选择单选按钮 ======================
# 在选择下拉菜单后，自动查找并选择单选按钮

# 自动查找页面上的单选按钮组
radios_groups = find_all_radios_groups()

# 如果找到单选按钮组，自动选择选项（根据实际需求修改）
if radios_groups:
    print(f"检测到 {len(radios_groups)} 个单选题，开始自动选择...")

    # 示例:处理多个单选选项，有几个单选，定义几个数值，第一个数值代表问题1，index = 1 ，代表第一个选项
    # [1,0]代表总共有两个问题，第一个问题选择第二项，第二个问题选择第一项。
    indices_to_select = [1, 0]
    for i, radio_group in enumerate(radios_groups):
        print(f"单选菜单 {i + 1}: 准备选择选项")
        if i < len(indices_to_select):
            select_radio_by_index(radio_group, indices_to_select[i])
        else:
            print("未定义索引，选择默认第一个选项")
            select_radio_by_index(radio_group, 0)  # 默认选择0

else:
    print("未检测到单选按钮组，跳过选择步骤")

# 单选按钮选择已完成，继续执行提交流程


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
print(N,"s后关闭页面")
for i in range(N):
    time.sleep(1)
    print(N-1-i,"s后关闭页面")
print(f"表单填写完成！")
print(f"执行耗时：{execution_time:.2f} 秒")
print(f"当前时间：{datetime.datetime.now()}")




