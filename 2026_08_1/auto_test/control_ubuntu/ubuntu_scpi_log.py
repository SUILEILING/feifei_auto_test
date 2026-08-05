#!/usr/bin/env python3
import time
import signal
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

FIREFOX_BINARY = "/snap/firefox/current/usr/lib/firefox/firefox"
GECKODRIVER_PATH = "/snap/bin/geckodriver"
URL = "http://127.0.0.1:8081/"

def create_driver():
    options = Options()
    options.add_argument("--window-size=1920,1080")
    options.binary_location = FIREFOX_BINARY
    service = Service(executable_path=GECKODRIVER_PATH)
    return webdriver.Firefox(service=service, options=options)

def get_scpi_panel_element(driver):
    try:
        element = driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div[2]/div[2]/div/div[2]")
        return element
    except:
        pass
    try:
        panels = driver.find_elements(By.CSS_SELECTOR, "div.topMessageBox")
        if len(panels) >= 2:
            return panels[1]
    except:
        pass
    try:
        label = driver.find_element(By.XPATH, "//*[contains(text(), 'Scpi Message')]")
        parent = label.find_element(By.XPATH, "./ancestor::div[contains(@class, 'topMessageBox')]")
        return parent
    except:
        pass
    return None

def get_scpi_clear_button(driver, scpi_panel):
    try:
        return scpi_panel.find_element(By.XPATH, ".//button[normalize-space()='Clear']")
    except:
        pass
    try:
        clear_btns = driver.find_elements(By.XPATH, "//button[normalize-space()='Clear']")
        if len(clear_btns) >= 2:
            return clear_btns[1]
    except:
        pass
    return None

def main():
    driver = create_driver()
    driver.get(URL)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    time.sleep(2)

    msg_btn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Message']"))
    )
    driver.execute_script("arguments[0].click();", msg_btn)
    print("[OK] Message 面板已打开")
    time.sleep(1)

    scpi_panel = get_scpi_panel_element(driver)
    if not scpi_panel:
        print("[FAIL] 未找到 SCPI 面板，请检查定位策略")
        driver.quit()
        return

    clear_btn = get_scpi_clear_button(driver, scpi_panel)
    if clear_btn:
        clear_btn.click()
        print("[OK] 已清空 SCPI 面板历史")
        time.sleep(0.5)
    else:
        print("[WARN] 未找到 Clear 按钮，可能无法清空历史")

    print("开始监控 SCPI 面板内容（按 Ctrl+C 停止并输出完整内容）...")
    last_text = ""
    try:
        while True:
            try:
                current_panel = get_scpi_panel_element(driver)
                if current_panel is None:
                    print("[WARN] 面板元素丢失，尝试重新获取...")
                    time.sleep(0.5)
                    continue
                text = current_panel.text
            except Exception as e:
                print(f"[WARN] 读取文本失败: {e}")
                time.sleep(0.5)
                continue

            if text != last_text:
                old_lines = last_text.splitlines()
                new_lines = text.splitlines()
                added = [line for line in new_lines if line not in old_lines]
                if added:
                    print(f"\n[{time.strftime('%Y-%m-%d %H:%M:%S')}] 新增 {len(added)} 行:")
                    for line in added[-10:]: 
                        print(f"  {line[:200]}")
                last_text = text
            time.sleep(0.5)
    except KeyboardInterrupt:
        print("\n[!] 用户中断")
        print("\n=== 当前 SCPI 面板完整内容 ===")
        print(last_text)
        with open("/tmp/scpi_panel_full.txt", "w", encoding="utf-8") as f:
            f.write(last_text)
        print("\n完整内容已保存到 /tmp/scpi_panel_full.txt")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()