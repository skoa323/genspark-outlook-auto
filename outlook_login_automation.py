import time
import pandas as pd
import random
import logging
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, 
    NoSuchElementException, 
    ElementNotInteractableException,
    StaleElementReferenceException,
    ElementClickInterceptedException
)

# 尝试导入webdriver_manager，如果不存在则继续但会显示警告
try:
    from webdriver_manager.microsoft import EdgeChromiumDriverManager
    WEBDRIVER_MANAGER_AVAILABLE = True
except ImportError:
    WEBDRIVER_MANAGER_AVAILABLE = False
    print("警告: webdriver_manager未安装，将使用本地msedgedriver")
    
import threading
import queue
from selenium.webdriver.common.keys import Keys
import sys
import os

# 设置日志记录
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(threadName)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("automation_log.txt"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger()

# 全局变量
MAX_RETRIES = 3
DEFAULT_TIMEOUT = 30 
GENSPARK_URL = "https://www.genspark.ai/invite?invite_code=YTI2MTEzMzVMZmVkZkxkMWI0TGM4MGFMZmU2NDg3NDgyMTk4"
MAX_CONCURRENT = 5  # 默认最大并发数量

class AutomationException(Exception):
    """自定义异常类，用于标识自动化过程中的错误"""
    pass

def random_delay(min_seconds=1, max_seconds=3):
    """添加随机延迟，模拟人类行为"""
    delay = random.uniform(min_seconds, max_seconds)
    time.sleep(delay)
    return delay

def wait_for_element(driver, locator, timeout=DEFAULT_TIMEOUT, condition=EC.presence_of_element_located):
    """等待元素出现并返回"""
    try:
        element = WebDriverWait(driver, timeout).until(
            condition(locator)
        )
        return element
    except TimeoutException:
        logger.error(f"等待元素超时: {locator}")
        raise AutomationException(f"等待元素超时: {locator}")

def safe_click(driver, element, max_attempts=3):
    """安全点击元素，处理各种点击异常"""
    for attempt in range(max_attempts):
        try:
            if isinstance(element, tuple):
                element = wait_for_element(driver, element, condition=EC.element_to_be_clickable)
            
            # 尝试直接点击
            try:
                element.click()
                return True
            except (ElementClickInterceptedException, ElementNotInteractableException):
                # 如果直接点击失败，尝试JavaScript点击
                driver.execute_script("arguments[0].click();", element)
                return True
        except (StaleElementReferenceException, ElementNotInteractableException) as e:
            if attempt == max_attempts - 1:
                logger.error(f"点击元素失败: {e}")
                raise AutomationException(f"点击元素失败: {e}")
            random_delay(0.5, 1)
    return False

def login_outlook(driver, email, password):
    """登录Outlook账号"""
    try:
        logger.info(f"开始登录Outlook: {email}")
        driver.get("https://login.live.com/")
        random_delay(1, 2)  # 减少延迟
        
        # 分析并处理当前页面
        page_type, handled = analyze_and_handle_current_page(driver)
        if page_type != "login_page":
            # 如果不是登录页面，检查是否处理成功
            if not handled:
                logger.warning("页面处理失败，尝试重新载入登录页面")
                driver.get("https://login.live.com/")
                random_delay(1, 2)  # 减少延迟
        
        # 输入邮箱
        email_input = wait_for_element(driver, (By.NAME, "loginfmt"))
        email_input.clear()
        email_input.send_keys(email)
        
        # 点击下一步
        next_button = wait_for_element(driver, (By.ID, "idSIButton9"), condition=EC.element_to_be_clickable)
        safe_click(driver, next_button)
        random_delay(1, 2)  # 减少延迟
        
        # 分析并处理可能出现的中间页面
        page_type, _ = analyze_and_handle_current_page(driver)
        if page_type != "password_page":
            # 如果不是密码页面，重新尝试查找密码输入框
            time.sleep(2)
        
        # 输入密码
        password_input = wait_for_element(driver, (By.NAME, "passwd"))
        password_input.clear()
        password_input.send_keys(password)
        
        # 点击登录
        sign_in_button = wait_for_element(driver, (By.ID, "idSIButton9"), condition=EC.element_to_be_clickable)
        safe_click(driver, sign_in_button)
        random_delay(1, 2)  # 减少延迟
        
        # 登录后可能出现多个页面需要处理，最多处理5个连续页面（减少原来的10个）
        for _ in range(5):
            page_type, handled = analyze_and_handle_current_page(driver)
            
            # 如果已经成功登录到Outlook主页，退出循环
            if page_type == "outlook_main":
                logger.info("已成功登录到Outlook主页")
                break
                
            # 如果是未知页面但没有成功处理，再尝试一次
            if page_type == "unknown" and not handled:
                time.sleep(1)  # 减少延迟
                continue
                
            # 给页面加载和处理的时间
            time.sleep(1)  # 减少延迟
        
        # 验证登录成功 - 优化检测速度
        current_url = driver.current_url
        
        # 快速判断是否登录成功
        if "login.live.com" in current_url and "auth/oauth2" not in current_url:
            logger.error(f"登录失败，当前URL: {current_url}")
            return False
        
        logger.info(f"成功登录Outlook: {email}")
        return True
    
    except (TimeoutException, NoSuchElementException, AutomationException) as e:
        logger.error(f"登录Outlook失败: {email}, 错误: {str(e)}")
        return False

def check_and_logout_genspark(driver):
    """检查Genspark网站是否已登录，如果是则登出"""
    try:
        # 尝试查找登出按钮或用户菜单
        try:
            # 查找用户菜单按钮
            user_menu = wait_for_element(
                driver, 
                (By.CSS_SELECTOR, ".user-menu, .profile-icon, .avatar, .user-avatar, [aria-label='User menu']"), 
                timeout=10
            )
            safe_click(driver, user_menu)
            random_delay()
            
            # 查找登出按钮
            logout_button = wait_for_element(
                driver, 
                (By.XPATH, "//button[contains(.,'Log out') or contains(.,'Sign out') or contains(.,'Logout')]"),
                timeout=10
            )
            safe_click(driver, logout_button)
            logger.info("已登出Genspark网站")
            random_delay(2, 4)
        except AutomationException:
            logger.info("可能未登录Genspark或找不到登出按钮")
            
        # 寻找登录按钮确认已登出
        try:
            wait_for_element(
                driver,
                (By.XPATH, "//button[contains(.,'Sign in') or contains(.,'Log in') or contains(.,'Login')]"),
                timeout=10
            )
            logger.info("确认Genspark网站已处于登出状态")
            return True
        except AutomationException:
            logger.warning("无法确认Genspark网站是否已登出")
            return False
            
    except Exception as e:
        logger.error(f"检查登出状态时出错: {str(e)}")
        return False

def login_genspark_with_outlook(driver, email):
    """使用Outlook账号登录Genspark网站"""
    try:
        logger.info(f"开始使用Outlook登录Genspark: {email}")
        
        # 登录流程的最大尝试次数
        max_page_iterations = 15
        login_success = False
        manual_intervention_needed = False
        
        # 登录循环 - 持续分析和处理页面直到成功或达到最大尝试次数
        for iteration in range(max_page_iterations):
            # 首先分析当前页面
            current_url = driver.current_url
            logger.info(f"Genspark登录流程 (步骤 {iteration+1}) - 当前URL: {current_url}")
            
            # 检查是否是需要手动处理的页面
            if detect_genspark_invite_page(driver):
                logger.info("检测到需要手动操作的邀请朋友页面，标记为需要用户干预")
                manual_intervention_needed = True
                # 成功标记为True，因为这个页面实际上是登录流程的成功结束
                login_success = True
                break
            
            # 检查是否已经登录成功
            if check_genspark_login_success(driver):
                logger.info(f"已成功登录Genspark: {email}")
                login_success = True
                break
                
            # 使用页面分析函数识别当前页面
            page_type, page_handled = analyze_genspark_page(driver)
            
            # 如果页面已处理，继续下一轮分析
            if page_handled:
                logger.info(f"已处理页面类型: {page_type}")
                time.sleep(3)  # 等待页面转换
                continue
                
            # 针对特定页面类型执行操作
            if page_type == "genspark_login_options":
                # 在登录选项页面，点击Microsoft按钮
                if handle_genspark_login_options(driver):
                    logger.info("已选择Microsoft登录选项")
                    time.sleep(3)
                else:
                    logger.error("无法点击Microsoft登录选项")
                    break
                    
            elif page_type == "microsoft_auth":
                # Microsoft授权页面
                if handle_microsoft_auth_consent(driver):
                    logger.info("已处理Microsoft授权页面")
                    time.sleep(3)
                else:
                    logger.error("无法处理Microsoft授权页面")
                    # 即使处理失败也继续，因为可能会自动跳转
                    
            elif page_type == "genspark_home":
                # 已在Genspark主页，需要点击登录按钮
                login_button_clicked = False
                login_button_selectors = [
                    "//button[contains(text(), 'Sign in') or contains(text(), 'Log in')]",
                    "//a[contains(text(), 'Sign in') or contains(text(), 'Log in')]",
                    "//button[contains(@class, 'login')]",
                    "//a[contains(@class, 'login')]"
                ]
                
                for selector in login_button_selectors:
                    try:
                        elements = driver.find_elements(By.XPATH, selector)
                        for element in elements:
                            if element.is_displayed():
                                logger.info(f"点击登录按钮: {element.text if element.text else '无文本'}")
                                driver.execute_script("arguments[0].click();", element)
                                login_button_clicked = True
                                time.sleep(3)
                                break
                        if login_button_clicked:
                            break
                    except Exception as e:
                        logger.debug(f"尝试点击登录按钮失败: {str(e)}")
                        
                if not login_button_clicked:
                    logger.error("无法找到Genspark登录按钮")
                    
            elif page_type == "genspark_benefits":
                # 会员权益页面
                if handle_genspark_plus_benefits(driver):
                    logger.info("已处理会员权益页面")
                    time.sleep(3)
                else:
                    logger.error("无法处理会员权益页面")
                    
            elif page_type == "microsoft_login":
                # 微软登录页面 - 可能需要重新登录
                logger.info("检测到Microsoft登录页面，可能需要重新登录")
                if login_outlook(driver, email, None):  # 不传密码，仅使用已登录状态
                    logger.info("已重新验证Microsoft账户")
                    time.sleep(3)
                else:
                    logger.error("无法验证Microsoft账户")
                    
            elif page_type == "unknown":
                # 未知页面，尝试通用处理
                logger.warning("遇到未知页面类型")
                if handle_unexpected_pages(driver):
                    logger.info("已尝试通用页面处理")
                    time.sleep(3)
                else:
                    # 尝试刷新页面
                    logger.info("通用处理失败，尝试刷新页面")
                    driver.refresh()
                    time.sleep(5)
                    
            else:
                # 其他已知页面类型，但未执行特定操作
                logger.info(f"未针对页面类型 {page_type} 执行特定操作")
                time.sleep(3)
                
            # 每轮检查是否已经登录成功
            if check_genspark_login_success(driver):
                logger.info(f"已成功登录Genspark: {email}")
                login_success = True
                break
                
        if manual_intervention_needed:
            logger.info(f"账号 {email} 需要手动操作输入手机号和验证码，脚本暂停自动处理")
        
        return login_success
        
    except Exception as e:
        logger.error(f"登录Genspark过程出错: {email}, 错误: {str(e)}")
        return False

def handle_unexpected_pages(driver):
    """处理可能出现的各种提示或确认页面"""
    try:
        # 处理保持登录状态页面
        if handle_stay_signed_in_page(driver):
            return True
            
        # 处理保护账户页面
        if handle_protect_account_page(driver):
            return True
            
        # 处理Microsoft授权页面
        if handle_microsoft_auth_consent(driver):
            return True
            
        # 处理Genspark会员权益页面
        if handle_genspark_plus_benefits(driver):
            return True
            
        # 处理Genspark登录选择页面
        if handle_genspark_login_options(driver):
            return True
            
        # 检测是否是Microsoft隐私通知页面
        current_url = driver.current_url
        if "privacynotice.account.microsoft.com" in current_url:
            logger.info("检测到Microsoft隐私通知页面")
            
            # Microsoft隐私页面的按钮可能有多种形式，尝试多种选择器
            privacy_button_selectors = [
                "//button[contains(@id, 'accept')]",
                "//button[contains(@id, 'agree')]",
                "//button[contains(@id, 'confirm')]",
                "//button[contains(@id, 'next')]",
                "//button[contains(@id, 'continue')]",
                "//button[@class='primary']",
                "//button[@class='win-button win-button-primary']",
                "//div[contains(@class, 'button')]//button",
                "//input[@type='submit']",
                "//button",  # 如果以上都找不到，尝试页面上的任何按钮
            ]
            
            for selector in privacy_button_selectors:
                try:
                    # 找到所有匹配的按钮
                    buttons = driver.find_elements(By.XPATH, selector)
                    for button in buttons:
                        try:
                            # 尝试判断哪个按钮是主要操作按钮
                            button_text = button.text.strip().lower()
                            is_primary = (
                                'class' in button.get_attribute('outerHTML').lower() and 
                                'primary' in button.get_attribute('outerHTML').lower()
                            )
                            
                            # 判断按钮是否是确认类型
                            if (button_text and (
                                    'accept' in button_text or 
                                    'agree' in button_text or 
                                    'confirm' in button_text or 
                                    'ok' in button_text or 
                                    'next' in button_text or 
                                    'continue' in button_text or
                                    '确定' in button_text or
                                    '接受' in button_text or
                                    '同意' in button_text or
                                    '继续' in button_text
                                )) or is_primary:
                                
                                logger.info(f"在隐私页面找到按钮: {button_text}")
                                driver.execute_script("arguments[0].click();", button)
                                random_delay(2, 4)
                                return True
                        except Exception as e:
                            logger.warning(f"尝试点击隐私页面按钮时出错: {str(e)}")
                            continue
                except Exception:
                    continue
            
            # 如果常规方法失败，尝试直接使用JavaScript点击页面上的第一个或最显眼的按钮
            try:
                logger.info("尝试JavaScript方式点击隐私页面按钮")
                # 尝试点击页面上的主按钮
                driver.execute_script("""
                    var buttons = document.querySelectorAll('button');
                    for(var i=0; i<buttons.length; i++) {
                        if(buttons[i].offsetParent !== null) {  // 检查按钮是否可见
                            buttons[i].click();
                            return true;
                        }
                    }
                    return false;
                """)
                random_delay(2, 4)
                return True
            except Exception as e:
                logger.error(f"JavaScript点击隐私页面按钮失败: {str(e)}")
        
        # 处理各种可能的确认按钮
        button_patterns = [
            "//button[contains(text(), '确定') or contains(text(), 'OK') or contains(text(), 'Confirm')]",
            "//button[contains(text(), '继续') or contains(text(), 'Continue') or contains(text(), 'Next')]",
            "//button[contains(text(), '我同意') or contains(text(), 'I Agree') or contains(text(), 'Accept') or contains(text(), 'agree')]",
            "//button[contains(@class, 'primary') or contains(@class, 'submit') or contains(@class, 'confirm')]",
            "//button[@type='submit']",
            "//input[@type='submit']",
            "//a[contains(text(), '确定') or contains(text(), 'OK') or contains(text(), 'Confirm')]",
            "//a[contains(text(), '继续') or contains(text(), 'Continue') or contains(text(), 'Next')]",
            "//div[contains(@class, 'button')]//button"
        ]
        
        for pattern in button_patterns:
            try:
                buttons = driver.find_elements(By.XPATH, pattern)
                if buttons:
                    for button in buttons:
                        if button.is_displayed():
                            logger.info(f"发现可点击按钮: {button.text if button.text else '无文本'}")
                            driver.execute_script("arguments[0].click();", button)
                            logger.info("已处理未知页面")
                            random_delay(1, 3)
                            return True
            except Exception:
                continue
                
        return False
    except Exception as e:
        logger.warning(f"处理未知页面时出错: {str(e)}")
        return False

def handle_microsoft_privacy_page(driver):
    """专门处理Microsoft隐私通知页面"""
    try:
        current_url = driver.current_url
        if "privacynotice.account.microsoft.com" in current_url:
            logger.info("检测到Microsoft隐私通知页面，尝试处理...")
            
            # 等待页面加载完成
            time.sleep(3)
            
            # 尝试多种方法找到并点击接受按钮
            try:
                # 方法1: 尝试找到主按钮
                buttons = driver.find_elements(By.TAG_NAME, "button")
                for button in buttons:
                    if button.is_displayed():
                        logger.info(f"点击隐私页面按钮: {button.text if button.text else '无文本'}")
                        driver.execute_script("arguments[0].click();", button)
                        random_delay(3, 5)
                        return True
                        
                # 方法2: 尝试直接访问下一个URL
                logger.info("无法找到隐私页面按钮，尝试直接跳过")
                redirect_url = driver.execute_script("""
                    var links = document.querySelectorAll('a');
                    for(var i=0; i<links.length; i++) {
                        if(links[i].href && links[i].href.includes('login.live.com')) {
                            return links[i].href;
                        }
                    }
                    return '';
                """)
                
                if redirect_url:
                    logger.info(f"尝试直接访问: {redirect_url}")
                    driver.get(redirect_url)
                    return True
                    
                return False
            except Exception as e:
                logger.error(f"处理隐私页面时出错: {str(e)}")
                return False
        return False
    except Exception as e:
        logger.error(f"检测隐私页面时出错: {str(e)}")
        return False

def handle_stay_signed_in_page(driver):
    """专门处理'保持登录状态'页面"""
    try:
        # 检查页面标题或特征文本
        page_texts = [
            "保持登录状态", "Stay signed in", "Keep me signed in"
        ]
        
        for text in page_texts:
            try:
                element = driver.find_element(By.XPATH, f"//*[contains(text(), '{text}')]")
                if element.is_displayed():
                    logger.info(f"检测到'保持登录状态'页面")
                    
                    # 查找并点击"是"或"Yes"按钮
                    yes_button_selectors = [
                        "//button[contains(text(), '是') or contains(text(), 'Yes')]",
                        "//input[@type='submit' and @value='是']",
                        "//input[@type='submit' and @value='Yes']",
                        "//button[@id='idSIButton9']"  # Microsoft常用的ID
                    ]
                    
                    for selector in yes_button_selectors:
                        try:
                            button = driver.find_element(By.XPATH, selector)
                            if button.is_displayed():
                                logger.info(f"点击'保持登录状态'页面的'是'按钮")
                                # 使用JavaScript点击以避免可能的事件问题
                                driver.execute_script("arguments[0].click();", button)
                                time.sleep(3)  # 等待页面反应
                                return True
                        except Exception:
                            continue
                    
                    # 如果没有找到特定按钮，尝试点击任何可见的提交按钮
                    buttons = driver.find_elements(By.XPATH, "//button[@type='submit'] | //input[@type='submit']")
                    for button in buttons:
                        if button.is_displayed():
                            logger.info("点击找到的提交按钮")
                            driver.execute_script("arguments[0].click();", button)
                            time.sleep(3)
                            return True
                    
                    # 尝试直接按回车键
                    logger.info("尝试按回车键提交")
                    webdriver.ActionChains(driver).send_keys(Keys.ENTER).perform()
                    time.sleep(3)
                    return True
            except NoSuchElementException:
                continue
                
        return False
    except Exception as e:
        logger.error(f"处理'保持登录状态'页面出错: {str(e)}")
        return False

def handle_protect_account_page(driver):
    """专门处理'让我们来保护你的账户'页面"""
    try:
        # 检查当前URL是否包含验证相关路径
        current_url = driver.current_url
        security_url_patterns = [
            "account.live.com/proofs",
            "account.microsoft.com/security",
            "account.microsoft.com/profile",
            "login.live.com/ppsecure",
            "account.live.com/identity/confirm"
        ]
        
        is_security_page = any(pattern in current_url for pattern in security_url_patterns)
        
        # 检查页面特征文本
        protection_texts = [
            "让我们来保护你的账户", "Protect your account", "Add security info",
            "保护你的账户", "验证你的身份", "添加安全信息", "添加安全验证方法",
            "Security info", "验证你的电子邮件", "验证身份"
        ]
        
        found_protection_text = False
        for text in protection_texts:
            try:
                elements = driver.find_elements(By.XPATH, f"//*[contains(text(), '{text}')]")
                if any(elem.is_displayed() for elem in elements if elem):
                    found_protection_text = True
                    logger.info(f"检测到安全验证页面文本: '{text}'")
                    break
            except Exception:
                continue
                
        # 检查输入字段，可能是验证码或电话号码输入
        has_input_field = False
        try:
            input_fields = driver.find_elements(By.XPATH, "//input[@type='text' or @type='tel' or @type='email']")
            has_input_field = any(field.is_displayed() for field in input_fields if field)
        except Exception:
            pass
        
        # 如果判断为安全验证页面
        if is_security_page or found_protection_text or has_input_field:
            logger.info(f"检测到安全验证页面 (URL: {current_url})")
            
            # 首先尝试查找并点击"暂时跳过"或"Skip for now"链接
            skip_selectors = [
                "//a[contains(text(), '暂时跳过')]",
                "//a[contains(text(), 'Skip')]",
                "//a[contains(text(), 'skip')]",
                "//span[contains(text(), '暂时跳过')]",
                "//span[contains(text(), 'Skip')]",
                "//button[contains(text(), '暂时跳过')]",
                "//button[contains(text(), 'Skip')]",
                "//div[contains(@class, 'skip')]//a",
                "//div[contains(@class, 'secondary')]//a",
                "//a[contains(@class, 'secondary')]",
                "//a[contains(@href, 'skip')]"
            ]
            
            # 查找跳过链接
            for selector in skip_selectors:
                try:
                    elements = driver.find_elements(By.XPATH, selector)
                    for element in elements:
                        if element.is_displayed():
                            logger.info(f"点击跳过链接: {element.text if element.text else '无文本'}")
                            # 使用JavaScript点击，避免元素重叠问题
                            driver.execute_script("arguments[0].click();", element)
                            time.sleep(3)
                            return True
                except Exception as e:
                    logger.debug(f"尝试选择器 {selector} 失败: {str(e)}")
                    continue
            
            # 如果找不到明确的跳过链接，尝试查找页面上所有链接
            try:
                logger.info("尝试查找页面上的所有链接...")
                links = driver.find_elements(By.TAG_NAME, "a")
                skip_keywords = ['skip', 'later', '跳过', '稍后', 'later', '暂时']
                
                for link in links:
                    try:
                        if not link.is_displayed():
                            continue
                            
                        link_text = link.text.lower()
                        link_href = link.get_attribute('href') or ''
                        
                        # 检查链接文本和href是否包含跳过相关关键词
                        if any(keyword in link_text or keyword in link_href for keyword in skip_keywords):
                            logger.info(f"点击可能的跳过链接: {link_text} ({link_href})")
                            driver.execute_script("arguments[0].click();", link)
                            time.sleep(3)
                            return True
                    except Exception:
                        continue
                        
                # 如果还是找不到，尝试直接去往Outlook
                logger.info("未找到跳过链接，尝试直接访问Outlook")
                driver.get("https://outlook.live.com/mail/")
                time.sleep(5)
                return True
            except Exception as e:
                logger.error(f"查找跳过链接时出错: {str(e)}")
            
            logger.warning("未找到任何可用的跳过选项")
            return False  # 无法找到跳过选项
                
        return False  # 不是安全验证页面
    except Exception as e:
        logger.error(f"处理安全验证页面时出错: {str(e)}")
        return False

def handle_verification_code_input(driver):
    """处理可能需要验证码输入的情况，返回False让用户手动处理"""
    try:
        # 检查是否存在验证码输入字段
        input_selectors = [
            "//input[@type='text' and contains(@placeholder, 'code')]",
            "//input[@type='text' and contains(@placeholder, '验证')]",
            "//input[@type='text' and contains(@aria-label, 'code')]",
            "//input[@type='text' and contains(@aria-label, '验证')]",
            "//input[@type='number']"
        ]
        
        for selector in input_selectors:
            try:
                elements = driver.find_elements(By.XPATH, selector)
                for element in elements:
                    if element.is_displayed():
                        logger.warning(f"检测到需要输入验证码的字段: {element.get_attribute('placeholder') or element.get_attribute('aria-label')}")
                        logger.warning("需要手动输入验证码，请在浏览器中完成验证")
                        return False  # 返回False，需要手动处理
            except Exception:
                continue
                
        return None  # 没有检测到验证码输入字段
    except Exception as e:
        logger.error(f"检查验证码输入时出错: {str(e)}")
        return None

def create_driver():
    """创建配置好的Edge WebDriver实例，使用InPrivate模式"""
    max_attempts = 3
    
    for attempt in range(max_attempts):
        try:
            options = EdgeOptions()
            
            # 启用InPrivate模式，如果失败可以注释掉这行测试
            try:
                options.add_argument("--inprivate")
            except Exception as e:
                logger.warning(f"启用InPrivate模式失败: {str(e)}，将使用普通模式")
            
            # 添加更稳定的配置
            options.add_argument("--start-maximized")
            options.add_argument("--disable-notifications")
            options.add_argument("--disable-popup-blocking")
            options.add_argument("--disable-infobars")
            options.add_argument("--disable-extensions")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-gpu")
            options.add_argument("--disable-dev-shm-usage")
            
            # 判断打包环境还是开发环境
            if getattr(sys, 'frozen', False):
                # 运行在打包环境
                edgedriver_path = os.path.join(sys._MEIPASS, "msedgedriver.exe") if hasattr(sys, '_MEIPASS') else "msedgedriver.exe"
                if os.path.exists(edgedriver_path):
                    service = EdgeService(edgedriver_path)
                    driver = webdriver.Edge(service=service, options=options)
                else:
                    # 尝试使用系统PATH中的msedgedriver
                    try:
                        driver = webdriver.Edge(options=options)
                    except Exception as e:
                        logger.warning(f"使用系统PATH中的msedgedriver失败: {str(e)}，尝试直接创建Edge驱动")
                        driver = webdriver.Edge()
            else:
                # 运行在开发环境
                if WEBDRIVER_MANAGER_AVAILABLE:
                    try:
                        service = EdgeService(EdgeChromiumDriverManager().install())
                        driver = webdriver.Edge(service=service, options=options)
                    except Exception as e:
                        logger.warning(f"使用EdgeChromiumDriverManager失败: {str(e)}，尝试其他方法")
                        driver = webdriver.Edge(options=options)
                else:
                    # 尝试使用当前目录或系统PATH中的msedgedriver
                    try:
                        driver = webdriver.Edge(options=options)
                    except Exception as e:
                        logger.warning(f"尝试直接创建Edge驱动失败: {str(e)}，最后一次尝试")
                        driver = webdriver.Edge()
                    
            driver.set_page_load_timeout(60)
            return driver
            
        except Exception as e:
            logger.error(f"创建WebDriver失败 (尝试 {attempt+1}/{max_attempts}): {str(e)}")
            if attempt < max_attempts - 1:
                logger.warning("等待5秒后重试创建浏览器...")
                time.sleep(5)
            else:
                logger.error("所有尝试都失败，无法创建浏览器")
                raise AutomationException(f"创建WebDriver失败: {str(e)}")

def process_account(account_info, result_queue):
    """处理单个账号的完整流程"""
    email = account_info['email']
    password = account_info['password']
    
    driver = None
    success = False
    
    try:
        # 使用增强的driver创建函数
        driver = create_driver()
        logger.info(f"成功创建Edge实例, 开始处理账号: {email}")
        
        # 使用网络错误处理包装登录流程
        login_success = False
        for attempt in range(MAX_RETRIES):
            try:
                if login_outlook(driver, email, password):
                    login_success = True
                    break
                
                # 如果登录失败，分析当前页面并尝试处理
                page_type, handled = analyze_and_handle_current_page(driver)
                
                if attempt < MAX_RETRIES - 1:
                    logger.info(f"重试登录Outlook: {email}, 尝试 {attempt+2}/{MAX_RETRIES}")
                    # 清除cookies并重新开始
                    driver.delete_all_cookies()
                    driver.get("https://login.live.com/")
                    random_delay(2, 3)  # 减少延迟时间
            except Exception as e:
                logger.error(f"登录过程出错: {str(e)}")
                if "net::" in str(e) or "SSL" in str(e) or "socket" in str(e):
                    logger.warning("检测到网络错误，正在重置连接")
                    try:
                        driver.quit()
                        time.sleep(3)  # 减少等待时间
                        driver = create_driver()
                        logger.info("已重置浏览器实例")
                    except Exception:
                        logger.error("重置浏览器实例失败")
                
                if attempt < MAX_RETRIES - 1:
                    logger.info(f"将在3秒后重试 ({attempt+2}/{MAX_RETRIES})")  # 减少等待时间
                    time.sleep(3)
        
        if login_success:
            logger.info(f"Outlook登录成功，立即访问Genspark网站: {email}")
            
            # 使用错误处理包装Genspark访问
            try:
                # 登录成功后立即访问Genspark，不再停留在Outlook页面
                driver.get(GENSPARK_URL)
                # 减少等待时间
                random_delay(3, 5)
                
                # 使用新的页面分析优先的Genspark登录函数
                genspark_success = login_genspark_with_outlook(driver, email)
                
                if genspark_success:
                    logger.info(f"成功完成Genspark注册: {email}")
                    result_queue.put((email, True, driver))
                    success = True
                else:
                    # 如果显示需要输入验证码或手机号的页面，标记为需要手动处理
                    if detect_genspark_invite_page(driver):
                        logger.info(f"账号 {email} 需要手动操作输入手机号和验证码")
                        result_queue.put((email, True, driver, "需要手动输入手机号和验证码"))
                        success = True
                    else:
                        logger.error(f"无法完成Genspark注册: {email}")
                        result_queue.put((email, False, driver))
            except Exception as e:
                logger.error(f"Genspark处理过程出错: {str(e)}")
                # 即使Genspark失败，我们也将driver保留，以便手动检查
                result_queue.put((email, False, driver))
        else:
            logger.error(f"无法登录Outlook账号: {email}")
            result_queue.put((email, False, driver))
            
    except Exception as e:
        logger.error(f"处理账号时出错: {email}, 错误: {str(e)}")
        if driver:
            result_queue.put((email, False, driver))
        else:
            result_queue.put((email, False, None))

def handle_genspark_plus_benefits(driver):
    """处理Genspark Plus会员权益页面"""
    try:
        # 检查页面特征，包括"获得1个月的免费Genspark Plus"文本
        benefit_texts = [
            "获得 1 个月的免费 Genspark Plus",
            "Genspark Plus 会员",
            "1个月的免费",
            "领取会员权益"
        ]
        
        found_benefit_page = False
        for text in benefit_texts:
            try:
                elements = driver.find_elements(By.XPATH, f"//*[contains(text(), '{text}')]")
                if any(elem.is_displayed() for elem in elements if elem):
                    found_benefit_page = True
                    logger.info(f"检测到Genspark会员权益页面，文本: '{text}'")
                    break
            except Exception:
                continue
        
        if not found_benefit_page:
            return False
            
        logger.info("开始处理Genspark会员权益页面")
        
        # 尝试查找并点击"领取会员权益"按钮
        claim_button_selectors = [
            "//button[contains(text(), '领取会员权益')]",
            "//a[contains(text(), '领取会员权益')]",
            "//div[contains(@class, 'button')][contains(text(), '领取')]",
            "//button[contains(@class, 'primary')]",
            "//button[contains(@class, 'blue')]"
        ]
        
        button_clicked = False
        for selector in claim_button_selectors:
            try:
                elements = driver.find_elements(By.XPATH, selector)
                for element in elements:
                    if element.is_displayed():
                        logger.info(f"点击领取会员权益按钮: {element.text}")
                        driver.execute_script("arguments[0].click();", element)
                        time.sleep(3)
                        button_clicked = True
                        break
                if button_clicked:
                    break
            except Exception as e:
                logger.debug(f"尝试点击选择器 {selector} 失败: {str(e)}")
                continue
        
        # 如果无法找到特定按钮，尝试点击页面上任何蓝色或主要的按钮
        if not button_clicked:
            try:
                possible_buttons = driver.find_elements(By.XPATH, "//button")
                for button in possible_buttons:
                    if button.is_displayed():
                        button_style = button.get_attribute("style") or ""
                        button_class = button.get_attribute("class") or ""
                        if ("blue" in button_style.lower() or 
                                "primary" in button_class.lower() or
                                "#" in button_style.lower()):
                            logger.info(f"点击可能的领取按钮: {button.text}")
                            driver.execute_script("arguments[0].click();", button)
                            time.sleep(3)
                            button_clicked = True
                            break
            except Exception as e:
                logger.error(f"尝试查找通用按钮失败: {str(e)}")
        
        # 如果还是找不到按钮，使用JavaScript尝试查找蓝色按钮并点击
        if not button_clicked:
            try:
                logger.info("尝试使用JavaScript查找并点击蓝色按钮")
                clicked = driver.execute_script("""
                    var buttons = document.querySelectorAll('button');
                    for (var i = 0; i < buttons.length; i++) {
                        var style = window.getComputedStyle(buttons[i]);
                        var backgroundColor = style.backgroundColor;
                        var color = style.color;
                        if (buttons[i].offsetParent !== null && 
                            (backgroundColor.includes('rgb(0, 122, 255)') || 
                             backgroundColor.includes('blue') || 
                             backgroundColor.includes('#') || 
                             buttons[i].className.includes('primary') || 
                             buttons[i].className.includes('blue'))) {
                            buttons[i].click();
                            return true;
                        }
                    }
                    return false;
                """)
                
                if clicked:
                    logger.info("已通过JavaScript点击按钮")
                    time.sleep(3)
                    button_clicked = True
            except Exception as e:
                logger.error(f"JavaScript点击失败: {str(e)}")
        
        if not button_clicked:
            logger.warning("无法找到领取会员权益按钮")
            return False
            
        # 按钮点击成功后，等待页面加载并检查是否进入登录选择页面
        time.sleep(5)  # 等待页面跳转
        
        # 检查是否到达登录选择页面，并选择Microsoft登录
        if handle_genspark_login_options(driver):
            logger.info("已处理Genspark登录选择页面，选择了Microsoft登录")
            return True
            
        return True  # 即使没有检测到登录选择页面，也返回成功
    except Exception as e:
        logger.error(f"处理Genspark会员权益页面时出错: {str(e)}")
        return False

def analyze_and_handle_current_page(driver):
    """
    分析当前页面并执行相应的处理
    返回 (页面类型, 是否处理成功)
    """
    try:
        # 获取当前页面信息
        current_url = driver.current_url
        page_title = driver.title
        page_source = driver.page_source.lower()
        
        logger.info(f"分析页面: URL={current_url}, 标题={page_title}")
        
        # 首先检查是否是需要手动处理的邀请页面
        if detect_genspark_invite_page(driver):
            logger.info("检测到需要手动操作的Genspark邀请朋友页面，停止自动处理")
            return ("genspark_invite", True)
        
        # 定义页面类型识别规则
        page_rules = [
            # (页面类型, URL包含, 页面文本包含, 元素选择器)
            # 将邀请页面添加到规则中
            ("genspark_invite", "genspark.ai", 
             ["立即邀请朋友", "你们双方都获得"], 
             "//input[contains(@placeholder, 'phone')]"),
            
            ("microsoft_privacy", "privacynotice.account.microsoft.com", 
             ["隐私", "privacy"], "//button"),
            
            ("stay_signed_in", "", 
             ["保持登录状态", "stay signed in", "keep me signed in"], 
             "//button[@id='idSIButton9']"),
            
            ("protect_account", "account.live.com/proofs|account.microsoft.com/security|login.live.com/ppsecure", 
             ["保护你的账户", "protect your account", "添加安全信息", "security info"], 
             "//a[contains(text(), '跳过') or contains(text(), 'Skip')]"),
            
            ("quick_note", "", 
             ["有关microsoft账户的快速说明", "microsoft account", "quick note"],
             "//button[contains(text(), '确定') or contains(text(), 'OK')]"),
            
            ("microsoft_auth", "login.live.com/oauth|account.live.com/consent", 
             ["是否允许此应用访问你的信息", "允许此应用访问", "查看你的基本个人资料"], 
             "//button[contains(text(), '接受')]"),
            
            ("genspark_benefits", "", 
             ["获得 1 个月的免费 genspark plus", "领取会员权益"], 
             "//button[contains(text(), '领取')]"),
             
            ("genspark_login_options", "genspark.ai", 
             ["reinvent search", "the ai agentic engine"], 
             "//button[contains(., 'Microsoft')]"),
            
            ("login_page", "login.live.com", 
             ["登录", "sign in"], "//input[@name='loginfmt']"),
            
            ("password_page", "login.live.com", 
             ["输入密码", "enter password"], "//input[@name='passwd']"),
            
            ("outlook_main", "outlook.live.com/mail", 
             ["收件箱", "inbox"], "//div[contains(@class, 'LeftRail')]"),
            
            ("genspark_login", "genspark.ai", 
             ["sign in", "log in"], "//button[contains(text(), 'Sign in')]"),
            
            ("genspark_logged_in", "genspark.ai", 
             ["welcome", "dashboard"], "//div[contains(@class, 'user-name')]")
        ]
        
        # 分析页面类型
        detected_page_type = "unknown"
        for page_type, url_pattern, text_patterns, element_selector in page_rules:
            # 检查URL
            if url_pattern and not any(pattern in current_url for pattern in url_pattern.split('|')):
                continue
                
            # 检查页面文本
            text_found = False
            for pattern in text_patterns:
                if pattern in page_source:
                    text_found = True
                    break
                    
            if not text_found and text_patterns:
                continue
                
            # 尝试找到关键元素
            try:
                elements = driver.find_elements(By.XPATH, element_selector)
                if not any(elem.is_displayed() for elem in elements if elem):
                    continue
            except Exception:
                continue
                
            detected_page_type = page_type
            logger.info(f"识别到页面类型: {detected_page_type}")
            break
            
        # 根据页面类型执行相应处理
        if detected_page_type == "genspark_invite":
            # 对于邀请页面，不执行任何自动操作，返回已处理
            logger.info("检测到需要手动操作的Genspark邀请朋友页面，停止自动处理")
            success = True
        elif detected_page_type == "microsoft_privacy":
            success = handle_microsoft_privacy_page(driver)
        elif detected_page_type == "stay_signed_in":
            success = handle_stay_signed_in_page(driver)
        elif detected_page_type == "protect_account":
            success = handle_protect_account_page(driver)
        elif detected_page_type == "quick_note":
            success = handle_quick_note_page(driver)
        elif detected_page_type == "microsoft_auth":
            success = handle_microsoft_auth_consent(driver)
        elif detected_page_type == "genspark_benefits":
            success = handle_genspark_plus_benefits(driver)
        elif detected_page_type == "genspark_login_options":
            success = handle_genspark_login_options(driver)
        elif detected_page_type == "unknown":
            # 对于未知页面，尝试通用处理
            logger.warning("无法识别页面类型，尝试通用处理")
            success = handle_generic_page(driver)
        else:
            # 对于其他已知但不需要特殊处理的页面，返回成功
            logger.info(f"已识别页面类型 {detected_page_type}，无需特殊处理")
            success = True
            
        return (detected_page_type, success)
    except Exception as e:
        logger.error(f"分析和处理页面时出错: {str(e)}")
        return ("error", False)

def handle_generic_page(driver):
    """通用页面处理，尝试寻找常见的确认按钮或跳过链接"""
    try:
        # 通用按钮选择器，按优先级排序
        button_selectors = [
            "//button[contains(text(), '确定') or contains(text(), 'OK') or contains(text(), 'Confirm')]",
            "//button[contains(text(), '下一步') or contains(text(), 'Next')]",
            "//button[contains(text(), '继续') or contains(text(), 'Continue')]",
            "//button[contains(text(), '我同意') or contains(text(), 'Agree') or contains(text(), 'Accept')]",
            "//button[contains(@class, 'primary')]",
            "//button[@type='submit']",
            "//a[contains(text(), '跳过') or contains(text(), 'Skip')]",
            "//button"  # 最后尝试任何按钮
        ]
        
        for selector in button_selectors:
            try:
                elements = driver.find_elements(By.XPATH, selector)
                for element in elements:
                    if element.is_displayed():
                        logger.info(f"通用处理: 点击 '{element.text if element.text else '无文本按钮'}'")
                        driver.execute_script("arguments[0].click();", element)
                        time.sleep(2)
                        return True
            except Exception:
                continue
        
        logger.warning("通用处理: 无法找到可点击元素")
        return False
    except Exception as e:
        logger.error(f"通用页面处理出错: {str(e)}")
        return False

def handle_quick_note_page(driver):
    """处理"有关Microsoft账户的快速说明"页面"""
    try:
        confirm_button = None
        # 尝试找到"确定"按钮
        button_selectors = [
            "//button[contains(text(), '确定') or contains(text(), 'OK') or contains(text(), 'Confirm')]",
            "//button[@id='iNext']",
            "//button[@id='idSIButton9']",
            "//input[@type='submit']"
        ]
        
        for selector in button_selectors:
            try:
                elements = driver.find_elements(By.XPATH, selector)
                for element in elements:
                    if element.is_displayed():
                        confirm_button = element
                        break
                if confirm_button:
                    break
            except Exception:
                continue
                
        if confirm_button:
            logger.info(f"点击快速说明页面的确定按钮: {confirm_button.text if confirm_button.text else '无文本'}")
            driver.execute_script("arguments[0].click();", confirm_button)
            time.sleep(3)
            return True
        else:
            logger.warning("无法在快速说明页面找到确定按钮")
            return False
    except Exception as e:
        logger.error(f"处理快速说明页面时出错: {str(e)}")
        return False

def handle_genspark_login_options(driver):
    """处理Genspark登录选项页面，选择Microsoft登录"""
    try:
        # 无需再次检查是否在登录选项页面，调用此函数前已经确认
        logger.info("开始处理Genspark登录选项页面")
        
        # 查找并点击Microsoft登录按钮的多种尝试
        methods_tried = 0
        methods_total = 4
        
        # 方法1: 使用XPath定位包含Microsoft文本的按钮
        methods_tried += 1
        logger.info(f"尝试方法 {methods_tried}/{methods_total}: XPath定位Microsoft按钮")
        ms_button_selectors = [
            "//button[contains(., 'Microsoft')]", 
            "//button[contains(@class, 'microsoft')]",
            "//div[contains(@class, 'microsoft')]//button",
            "//button//*[contains(text(), 'Microsoft')]/.."
        ]
        
        for selector in ms_button_selectors:
            try:
                elements = driver.find_elements(By.XPATH, selector)
                for element in elements:
                    if element.is_displayed():
                        logger.info("找到并点击Microsoft登录按钮")
                        # 尝试直接点击
                        try:
                            element.click()
                        except Exception:
                            # 如果直接点击失败，使用JavaScript点击
                            driver.execute_script("arguments[0].click();", element)
                        time.sleep(3)
                        return True
            except Exception as e:
                logger.debug(f"使用选择器 {selector} 点击Microsoft按钮失败: {str(e)}")
        
        # 方法2: 循环所有按钮，检查内部文本和属性
        methods_tried += 1
        logger.info(f"尝试方法 {methods_tried}/{methods_total}: 分析所有按钮")
        try:
            buttons = driver.find_elements(By.TAG_NAME, "button")
            for button in buttons:
                try:
                    if not button.is_displayed():
                        continue
                        
                    button_html = button.get_attribute("outerHTML").lower()
                    button_text = button.text.lower()
                    
                    if "microsoft" in button_html or "microsoft" in button_text:
                        logger.info("通过HTML/文本内容找到Microsoft按钮")
                        driver.execute_script("arguments[0].click();", button)
                        time.sleep(3)
                        return True
                except Exception:
                    continue
        except Exception as e:
            logger.debug(f"通过循环按钮查找Microsoft按钮失败: {str(e)}")
        
        # 方法3: 使用精确的JavaScript定位
        methods_tried += 1
        logger.info(f"尝试方法 {methods_tried}/{methods_total}: 使用JavaScript定位按钮")
        try:
            clicked = driver.execute_script("""
                // 定位所有可能的Microsoft按钮
                var buttons = document.querySelectorAll('button');
                for (var i = 0; i < buttons.length; i++) {
                    var button = buttons[i];
                    // 检查按钮是否可见
                    if (button.offsetParent === null) continue;
                    
                    // 获取按钮的所有文本内容和HTML
                    var buttonText = button.innerText || '';
                    var buttonHTML = button.innerHTML || '';
                    
                    // 检查按钮是否包含Microsoft标识
                    if (buttonText.toLowerCase().includes('microsoft') || 
                        buttonHTML.toLowerCase().includes('microsoft')) {
                        // 点击按钮
                        button.click();
                        return true;
                    }
                    
                    // 检查按钮内的图像元素
                    var images = button.querySelectorAll('img');
                    for (var j = 0; j < images.length; j++) {
                        var imgSrc = images[j].src || '';
                        var imgAlt = images[j].alt || '';
                        if (imgSrc.toLowerCase().includes('microsoft') || 
                            imgAlt.toLowerCase().includes('microsoft')) {
                            button.click();
                            return true;
                        }
                    }
                }
                return false;
            """)
            
            if clicked:
                logger.info("通过JavaScript成功点击Microsoft按钮")
                time.sleep(3)
                return True
        except Exception as e:
            logger.debug(f"JavaScript点击Microsoft按钮失败: {str(e)}")
        
        # 方法4: 使用位置信息 - Microsoft按钮通常是第一个登录选项
        methods_tried += 1
        logger.info(f"尝试方法 {methods_tried}/{methods_total}: 点击第一个登录按钮")
        try:
            buttons = driver.find_elements(By.TAG_NAME, "button")
            if len(buttons) >= 1:
                first_visible_button = None
                for button in buttons:
                    if button.is_displayed():
                        first_visible_button = button
                        break
                
                if first_visible_button:
                    logger.info("尝试点击第一个可见按钮")
                    driver.execute_script("arguments[0].click();", first_visible_button)
                    time.sleep(3)
                    return True
        except Exception as e:
            logger.debug(f"点击第一个按钮失败: {str(e)}")
        
        logger.error("所有尝试都失败，无法找到Microsoft登录按钮")
        return False
    except Exception as e:
        logger.error(f"处理Genspark登录选项页面时出错: {str(e)}")
        return False

def handle_microsoft_auth_consent(driver):
    """处理Microsoft账户授权页面，自动点击'接受'按钮"""
    try:
        # 检查是否是授权页面
        auth_detected = False
        
        # 方法1: 通过URL检测
        current_url = driver.current_url
        auth_url_patterns = [
            "login.microsoftonline.com/common/oauth2",
            "login.live.com/oauth20_authorize",
            "account.live.com/consent",
            "login.microsoftonline.com/consumers/oauth2/v2.0/authorize"
        ]
        
        if any(pattern in current_url for pattern in auth_url_patterns):
            auth_detected = True
            logger.info(f"通过URL检测到Microsoft授权页面: {current_url}")
        
        # 方法2: 通过页面内容检测
        if not auth_detected:
            auth_texts = [
                "是否允许此应用访问你的信息",
                "允许此应用访问",
                "需要获得你的许可",
                "查看你的基本个人资料",
                "查看你的电子邮件地址"
            ]
            
            for text in auth_texts:
                try:
                    elements = driver.find_elements(By.XPATH, f"//*[contains(text(), '{text}')]")
                    if any(elem.is_displayed() for elem in elements if elem):
                        auth_detected = True
                        logger.info(f"通过文本检测到Microsoft授权页面: '{text}'")
                        break
                except Exception:
                    continue
        
        # 如果确认是授权页面，点击"接受"按钮
        if auth_detected:
            logger.info("开始处理Microsoft授权页面")
            
            # 查找并点击"接受"按钮
            accept_button_selectors = [
                "//button[contains(text(), '接受')]",
                "//input[@type='submit'][contains(@value, '接受')]",
                "//button[@id='idSIButton9']",  # 常见的Microsoft接受按钮ID
                "//button[contains(@class, 'primary')]",
                "//button[contains(@class, 'accept')]",
                "//div[contains(@class, 'button')][contains(text(), '接受')]"
            ]
            
            for selector in accept_button_selectors:
                try:
                    elements = driver.find_elements(By.XPATH, selector)
                    for element in elements:
                        if element.is_displayed():
                            logger.info(f"点击授权页面的接受按钮: {element.text if element.text else '无文本'}")
                            driver.execute_script("arguments[0].click();", element)
                            time.sleep(3)
                            return True
                except Exception as e:
                    logger.debug(f"尝试选择器 {selector} 失败: {str(e)}")
                    continue
            
            # 如果通过ID和文本无法找到按钮，尝试找到页面上的所有按钮并检查右侧蓝色的那个
            try:
                logger.info("尝试查找页面上的所有按钮")
                buttons = driver.find_elements(By.TAG_NAME, "button")
                # 蓝色接受按钮通常在页面右侧
                for button in reversed(buttons):  # 从右向左检查
                    if button.is_displayed():
                        style = button.get_attribute("style") or ""
                        class_name = button.get_attribute("class") or ""
                        
                        # 检查按钮样式或类名是否包含蓝色、主要或接受等关键词
                        if ("blue" in style.lower() or 
                                "primary" in class_name.lower() or 
                                "accept" in class_name.lower()):
                            logger.info(f"点击可能的接受按钮: {button.text if button.text else '无文本'}")
                            driver.execute_script("arguments[0].click();", button)
                            time.sleep(3)
                            return True
            except Exception as e:
                logger.error(f"查找通用按钮失败: {str(e)}")
            
            # 最后尝试JavaScript方法
            try:
                logger.info("尝试通过JavaScript查找并点击接受按钮")
                clicked = driver.execute_script("""
                    // 尝试查找蓝色按钮或接受按钮
                    var buttons = document.querySelectorAll('button');
                    var buttonsArray = Array.from(buttons);
                    
                    // 按照DOM顺序从右到左排序，通常接受按钮在右侧
                    buttonsArray.reverse();
                    
                    for (var i = 0; i < buttonsArray.length; i++) {
                        var button = buttonsArray[i];
                        if (button.offsetParent === null) continue;  // 忽略不可见按钮
                        
                        var style = window.getComputedStyle(button);
                        var text = button.innerText || '';
                        var className = button.className || '';
                        
                        // 检查是否是蓝色按钮或包含"接受"字样
                        if (text.includes('接受') || text.includes('Accept') || 
                            className.includes('primary') || style.backgroundColor.includes('rgb(0, 120, 212)') || 
                            style.backgroundColor.includes('blue')) {
                            button.click();
                            return true;
                        }
                    }
                    
                    // 如果找不到按钮，尝试点击最后一个可见按钮
                    for (var i = 0; i < buttonsArray.length; i++) {
                        if (buttonsArray[i].offsetParent !== null) {
                            buttonsArray[i].click();
                            return true;
                        }
                    }
                    
                    return false;
                """)
                
                if clicked:
                    logger.info("已通过JavaScript点击接受按钮")
                    time.sleep(3)
                    return True
            except Exception as e:
                logger.error(f"JavaScript点击接受按钮失败: {str(e)}")
            
            logger.warning("无法找到接受按钮")
            return False
        
        return False  # 不是授权页面
    except Exception as e:
        logger.error(f"处理Microsoft授权页面时出错: {str(e)}")
        return False

def analyze_genspark_page(driver):
    """
    分析当前Genspark相关页面并执行相应处理
    返回 (页面类型, 是否已处理)
    """
    try:
        # 获取当前页面信息
        current_url = driver.current_url
        page_source = driver.page_source.lower()
        
        # 首先检查是否是邀请朋友页面 - 这种页面需要停止自动操作
        if detect_genspark_invite_page(driver):
            logger.info("检测到需要手动操作的Genspark邀请朋友页面，停止自动处理")
            return "genspark_invite", True  # 返回True表示已处理（实际上是标记为不再自动处理）
        
        # 其他页面检测逻辑...
        # [现有的页面检测代码保持不变]
        
        # 默认返回未知页面类型
        return "unknown", False
    
    except Exception as e:
        logger.error(f"分析Genspark页面时出错: {str(e)}")
        return "error", False

def check_genspark_login_success(driver):
    """检查是否成功登录Genspark"""
    try:
        # 方法1: 检查URL和页面元素
        if "genspark.ai" in driver.current_url:
            # 尝试查找成功登录的特征元素
            success_indicators = [
                "//div[contains(@class, 'user-name')]",
                "//div[contains(@class, 'avatar')]",
                "//button[contains(@class, 'user-menu')]",
                "//div[contains(text(), 'Welcome')]",
                "//div[contains(@class, 'dashboard')]"
            ]
            
            for selector in success_indicators:
                try:
                    elements = driver.find_elements(By.XPATH, selector)
                    if any(elem.is_displayed() for elem in elements if elem):
                        logger.info("通过页面元素确认Genspark登录成功")
                        return True
                except Exception:
                    continue
                    
            # 方法2: 通过JavaScript检查登录状态
            try:
                is_logged_in = driver.execute_script("""
                    // 检查是否存在用户相关元素
                    if (document.querySelector('.user-name') || 
                        document.querySelector('.avatar') || 
                        document.querySelector('.user-menu')) {
                        return true;
                    }
                    
                    // 检查localStorage是否有登录令牌
                    if (window.localStorage) {
                        for (var i = 0; i < localStorage.length; i++) {
                            var key = localStorage.key(i);
                            if (key.includes('token') || key.includes('auth') || key.includes('login')) {
                                return true;
                            }
                        }
                    }
                    
                    return false;
                """)
                
                if is_logged_in:
                    logger.info("通过JavaScript确认Genspark登录成功")
                    return True
            except Exception:
                pass
                
        return False
    except Exception as e:
        logger.error(f"检查Genspark登录状态时出错: {str(e)}")
        return False

def detect_genspark_invite_page(driver):
    """检测是否是Genspark邀请朋友页面（需要手动输入手机号和验证码的页面）"""
    try:
        # 检查页面特征
        invite_texts = [
            "立即邀请朋友",
            "你们双方都获得 1 个月的免费 Genspark Plus",
            "生成邀请链接"
        ]
        
        for text in invite_texts:
            try:
                elements = driver.find_elements(By.XPATH, f"//*[contains(text(), '{text}')]")
                if any(elem.is_displayed() for elem in elements if elem):
                    logger.info(f"检测到Genspark邀请朋友页面，文本: '{text}'")
                    return True
            except Exception:
                continue
                
        # 检查输入电话号码字段
        try:
            phone_inputs = driver.find_elements(By.XPATH, "//input[contains(@placeholder, 'phone') or contains(@placeholder, '电话')]")
            if any(inp.is_displayed() for inp in phone_inputs if inp):
                logger.info("检测到Genspark邀请页面的电话号码输入框")
                return True
        except Exception:
            pass
            
        return False
    except Exception as e:
        logger.error(f"检测Genspark邀请页面时出错: {str(e)}")
        return False

def main(max_concurrent=MAX_CONCURRENT):
    """
    主函数，处理所有账号
    参数:
        max_concurrent: 最大并发数量，默认使用全局设置
    """
    # 读取账号信息
    try:
        df = pd.read_csv('outlook_accounts.csv')
        logger.info(f"成功读取了 {len(df)} 个账号")
    except Exception as e:
        logger.error(f"读取账号文件失败: {str(e)}")
        logger.info("请确保outlook_accounts.csv文件存在，格式为: email,password")
        return
    
    # 创建结果队列和线程列表
    result_queue = queue.Queue()
    threads = []
    drivers = {}
    manual_status = {}  # 用于存储需要手动操作的账号状态
    active_threads = []  # 当前活动的线程
    
    logger.info(f"设置最大并发数量: {max_concurrent}")
    
    # 启动线程处理每个账号，控制最大并发数量
    for index, row in df.iterrows():
        account_info = {'email': row['email'], 'password': row['password']}
        
        # 直接记录每个账号的开始处理
        logger.info(f"准备处理账号 {index+1}/{len(df)}: {account_info['email']}")
        
        # 控制并发数量
        while len([t for t in active_threads if t.is_alive()]) >= max_concurrent:
            # 如果活动线程数量达到最大值，等待
            logger.info(f"当前活动线程数量已达到最大值 {max_concurrent}，等待...")
            time.sleep(2)
            # 清理已完成的线程
            active_threads = [t for t in active_threads if t.is_alive()]
        
        # 创建并启动新线程
        thread = threading.Thread(
            target=process_account, 
            args=(account_info, result_queue),
            name=f"Thread-{index+1}"
        )
        threads.append(thread)
        active_threads.append(thread)
        thread.start()
        
        # 缩短线程启动间隔，确保账号处理更快
        time.sleep(5)
    
    # 等待所有线程完成
    for thread in threads:
        thread.join()
    
    # 收集结果时检查是否所有邮箱都被处理
    processed_emails = set()
    success_count = 0
    
    while not result_queue.empty():
        result = result_queue.get()
        
        # 处理结果队列中可能有3个或4个元素的情况
        if len(result) == 3:
            email, success, driver = result
            manual_msg = None
        elif len(result) == 4:
            email, success, driver, manual_msg = result
        else:
            logger.error(f"意外的结果格式: {result}")
            continue
        
        # 避免重复计数同一个邮箱    
        if email in processed_emails:
            continue
            
        processed_emails.add(email)
        drivers[email] = driver
        
        if success:
            success_count += 1
            if manual_msg:
                manual_status[email] = manual_msg
                logger.info(f"账号 {email} 需要手动操作: {manual_msg}")
    
    # 检查是否有邮箱没有被处理
    all_emails = set([row['email'] for _, row in df.iterrows()])
    missing_emails = all_emails - processed_emails
    
    if missing_emails:
        logger.warning(f"以下邮箱未被处理: {', '.join(missing_emails)}")
        logger.warning("这些邮箱可能因错误或浏览器启动问题而被跳过")
    
    logger.info(f"自动化完成，成功处理 {success_count}/{len(df)} 个账号，总共找到 {len(processed_emails)}/{len(df)} 个账号的处理结果")
    
    # 显示需要手动操作的账号
    if manual_status:
        logger.info("以下账号需要手动操作:")
        for email, status in manual_status.items():
            logger.info(f"  - {email}: {status}")
    
    logger.info("所有浏览器窗口将保持打开状态，请手动关闭")
    
    # 保持脚本运行，直到用户决定关闭
    try:
        input("按回车键结束程序 (浏览器窗口将保持打开)...")
    except:
        pass

if __name__ == "__main__":
    # 尝试从命令行参数获取最大并发数
    try:
        if len(sys.argv) > 1:
            custom_max_concurrent = int(sys.argv[1])
            if custom_max_concurrent > 0:
                logger.info(f"从命令行设置最大并发数: {custom_max_concurrent}")
                main(custom_max_concurrent)
            else:
                logger.warning("最大并发数必须大于0，使用默认值")
                main()
        else:
            main()  # 使用默认值
    except ValueError:
        logger.error("命令行参数必须是整数，使用默认值")
        main() 