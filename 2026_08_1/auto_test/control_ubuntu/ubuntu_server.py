﻿import socket
import threading
import subprocess
import re
import time
import base64
import gzip
import os
import sys
import select
import signal
import hashlib
import tempfile
import json
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


HOST = '0.0.0.0'
PORT = 9999
FIREFOX_BINARY = "/snap/firefox/current/usr/lib/firefox/firefox"
GECKODRIVER_PATH = "/snap/bin/geckodriver"
URL = "http://127.0.0.1:8081/"
VERSION = "YC1100.1.00.03.10"
SUDO_PASSWORD = "yc"

INSTRUMENT_IP = "192.168.30.122"
INSTRUMENT_PORT = 5025


_log_thread = None
_log_stop_event = None
_log_file_path = None
_scpi_driver = None
_scpi_ready = False
_scpi_lines = []

def log_with_timestamp(line: str):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    stamped = f"[{timestamp}] {line}"
    print(stamped, flush=True)
    return stamped

def create_driver():
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--window-size=1920,1080")
    options.binary_location = FIREFOX_BINARY
    service = Service(executable_path=GECKODRIVER_PATH)
    options.add_argument("--disable-gpu")
    return webdriver.Firefox(service=service, options=options)

def get_timestamp():
    return datetime.now().strftime('%Y%m%d_%H%M%S')

def read_memory_info():
    try:
        info = {}
        with open('/proc/meminfo', 'r') as f:
            for line in f:
                parts = line.split(':')
                if len(parts) == 2:
                    val = parts[1].strip().split()
                    if val:
                        try:
                            info[parts[0].strip()] = int(val[0])  # 单位 kB
                        except ValueError:
                            pass
        def gib(kb):
            return kb / 1024 / 1024
        mem_total = info.get('MemTotal', 0)
        mem_avail = info.get('MemAvailable', 0)
        swap_total = info.get('SwapTotal', 0)
        swap_free = info.get('SwapFree', 0)
        used_pct = (1 - mem_avail / mem_total) * 100 if mem_total else 0
        swap_used = swap_total - swap_free
        swap_pct = (swap_used / swap_total * 100) if swap_total else 0
        return (f"MemTotal={gib(mem_total):.2f}G MemAvailable={gib(mem_avail):.2f}G "
                f"MemUsed={used_pct:.0f}% "
                f"Swap={gib(swap_used):.2f}/{gib(swap_total):.2f}G({swap_pct:.0f}%)")
    except Exception as e:
        return f"读取内存信息失败: {e}"

def _read_cpu_times():
    overall = (None, None)
    percore = {}
    try:
        with open('/proc/stat', 'r') as f:
            for line in f:
                if line.startswith('cpu'):
                    parts = line.split()
                    label = parts[0]
                    vals = [int(x) for x in parts[1:]]
                    idle = vals[3] + (vals[4] if len(vals) > 4 else 0)  # idle + iowait
                    if label == 'cpu':
                        overall = (sum(vals), idle)
                    elif label[3:].isdigit():
                        percore[label] = (sum(vals), idle)
    except Exception:
        pass
    return overall[0], overall[1], percore

def read_cpu_info():
    try:
        cores = os.cpu_count() or 1
        try:
            with open('/proc/loadavg', 'r') as f:
                load_str = "/".join(f.read().split()[:3])
        except Exception:
            load_str = "N/A"

        t1, i1, per1 = _read_cpu_times()
        time.sleep(0.15)
        t2, i2, per2 = _read_cpu_times()

        if None not in (t1, i1, t2, i2) and t2 != t1:
            util_str = f"{(1 - (i2 - i1) / (t2 - t1)) * 100:.0f}%"
        else:
            util_str = "N/A"

        def core_index(label):
            return int(label[3:])
        percore_parts = []
        for label in sorted(per1.keys(), key=core_index):
            if label not in per2:
                continue
            tot1, idl1 = per1[label]
            tot2, idl2 = per2[label]
            if tot2 == tot1:
                percore_parts.append(f"{label}=N/A")
                continue
            core_util = (1 - (idl2 - idl1) / (tot2 - tot1)) * 100
            percore_parts.append(f"{label}={core_util:.0f}%")
        percore_str = " ".join(percore_parts) if percore_parts else "N/A"

        return f"util={util_str} load={load_str} ({cores}cores) percore[{percore_str}]"
    except Exception as e:
        return f"读取CPU信息失败: {e}"

def check_oom(max_lines=200):
    pattern = re.compile(r'oom|out of memory|killed process', re.IGNORECASE)
    oom_lines = []
    try:
        proc = subprocess.run(
            ["journalctl", "-k", "--no-pager", "-n", str(max_lines)],
            capture_output=True, text=True, timeout=10
        )
        oom_lines = [l for l in proc.stdout.splitlines() if pattern.search(l)]
    except Exception:
        pass

    if not oom_lines:
        try:
            res = subprocess.run(
                f"echo '{SUDO_PASSWORD}' | sudo -S dmesg -T 2>/dev/null | tail -n {max_lines}",
                shell=True, capture_output=True, text=True, timeout=10
            )
            oom_lines = [l for l in res.stdout.splitlines() if pattern.search(l)]
        except Exception:
            pass

    if oom_lines:
        return f"⚠️ 检测到 {len(oom_lines)} 条 OOM 记录，最近: " + " || ".join(oom_lines[-3:])
    return "未检测到 OOM"

def report_resource(conn, stage):
    res_line = f"[RES][{stage}] Mem: {read_memory_info()} | CPU: {read_cpu_info()}"
    oom_line = f"[OOM][{stage}] {check_oom()}"
    log_with_timestamp(res_line)
    log_with_timestamp(oom_line)
    try:
        conn.sendall((res_line + "\n").encode('utf-8'))
        conn.sendall((oom_line + "\n").encode('utf-8'))
    except Exception:
        pass

def send_file_via_base64(conn, file_path, compress_threshold_mb=5):
    if not os.path.exists(file_path):
        conn.sendall(f"!FILE:{os.path.basename(file_path)}:FILE_NOT_FOUND\n".encode())
        log_with_timestamp(f"文件不存在: {file_path}")
        return
    try:
        file_size = os.path.getsize(file_path)
        compress_threshold = compress_threshold_mb * 1024 * 1024
        if file_size == 0:
            conn.sendall(f"!FILE:{os.path.basename(file_path)}:FILE_EMPTY\n".encode())
            return
        with open(file_path, 'rb') as f:
            file_data = f.read()
        if file_size > compress_threshold:
            compressed_data = gzip.compress(file_data, compresslevel=6)
            b64_str = base64.b64encode(compressed_data).decode('ascii')
            filename = os.path.basename(file_path)
            conn.sendall(f"!FILE:{filename}.gz:{b64_str}\n".encode())
            log_with_timestamp(f"已发送压缩文件: {filename}.gz ({len(compressed_data)} bytes)")
        else:
            b64_str = base64.b64encode(file_data).decode('ascii')
            conn.sendall(f"!FILE:{os.path.basename(file_path)}:{b64_str}\n".encode())
            log_with_timestamp(f"已发送文件: {os.path.basename(file_path)} ({len(file_data)} bytes)")
    except Exception as e:
        conn.sendall(f"!FILE:{os.path.basename(file_path)}:ERROR:{e}\n".encode())
        log_with_timestamp(f"发送文件失败 {file_path}: {e}")

def _send_scpi_command(cmd: str) -> bool:
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(5)
            s.connect((INSTRUMENT_IP, INSTRUMENT_PORT))
            s.sendall((cmd + "\n").encode("ascii"))
            log_with_timestamp(f"SCPI 命令已发送: {cmd}")
            return True
    except Exception as e:
        log_with_timestamp(f"发送 SCPI 命令失败: {e}")
        return False

def do_restart(conn):
    try:
        driver = create_driver()
        driver.get(URL)
        wait = WebDriverWait(driver, 20)
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        time.sleep(1)

        restart_xpath = "//div[text()='Restart' and contains(@class, 'headBtn')]"
        btn = driver.find_element(By.XPATH, restart_xpath)
        btn.click()
        log_with_timestamp("已点击 Restart")

        try:
            WebDriverWait(driver, 2).until(EC.alert_is_present())
            alert = driver.switch_to.alert
            alert.accept()
            log_with_timestamp("已处理 alert")
        except:
            pass

        time.sleep(1)

        try:
            popup = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'el-message-box')]"))
            )
            log_with_timestamp("检测到弹窗")

            ok_btn = popup.find_element(By.XPATH, ".//button[contains(@class, 'el-button--primary')]")
            ok_btn.click()
            log_with_timestamp("已点击 OK 确认按钮")
        except Exception as e:
            log_with_timestamp(f"处理确认弹窗失败: {e}")
            try:
                ok_btn = driver.find_element(By.XPATH, "//button[contains(translate(., 'OK', 'ok'), 'ok')]")
                ok_btn.click()
                log_with_timestamp("备用方式点击 OK")
            except:
                pass

        driver.quit()
        report_resource(conn, "*RST/重启后")
        conn.sendall("[OK] 重启完成\n".encode())
    except Exception as e:
        conn.sendall(f"[EXCEPTION] do_restart 失败: {e}\n".encode())


def do_configure_log_and_restart(conn, kwargs):
    """整体 flow:下发 CONFigure:VERSion:LOG:STATe → 下发 *RST → 调用原始 restart"""
    raw_params = kwargs.get('params', '')
    try:
        params = json.loads(base64.b64decode(raw_params).decode('utf-8')) if raw_params else {}
    except Exception as e:
        conn.sendall(f"[ERROR] 解析 params 失败: {e}\n".encode())
        return

    log_level = str(params.get('log_level', 'info')).strip().lower()
    checkbox_items = [f"{k},{int(v)}" for k, v in params.items() if k != 'log_level']
    sub_parts = [log_level] + checkbox_items
    scpi_body = ";".join(sub_parts)
    scpi_cmd = f'CONFigure:VERSion:LOG:STATe "{scpi_body}"'

    if _send_scpi_command(scpi_cmd):
        conn.sendall("[OK] 日志级别 SCPI 命令已下发\n".encode())
    else:
        conn.sendall("[WARN] SCPI 命令发送失败，尝试继续\n".encode())

    if _send_scpi_command("*RST"):
        conn.sendall("[OK] *RST 已下发\n".encode())
    else:
        conn.sendall("[WARN] *RST 发送失败，尝试继续重启\n".encode())
    time.sleep(5)

    do_restart(conn)


def collect_signal_logs(conn, duration_sec):
    try:
        driver = create_driver()
        driver.get(URL)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        time.sleep(1)
        msg_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Message']"))
        )
        driver.execute_script("arguments[0].click();", msg_btn)
        time.sleep(1)
        ts = get_timestamp()
        signal_file = f"/tmp/{ts}_signal_message_log.txt"
        seen_lines = set()
        final_log = []
        start_time = time.time()
        while (time.time() - start_time) < duration_sec:
            current_text = driver.execute_script("return document.body.innerText;")
            lines = current_text.split('\n')
            for line in lines:
                line = line.strip()
                if re.search(r'\d{4}-\d{2}-\d{2}', line) and 'Send->' not in line and 'Recv<-' not in line:
                    if line and line not in seen_lines:
                        seen_lines.add(line)
                        final_log.append(line)
            time.sleep(0.5)
        if final_log:
            timestamp_pattern = re.compile(r'(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}(?::\d{3})?)')
            def extract_sort_key(line):
                match = timestamp_pattern.search(line)
                if match:
                    ts_str = match.group(1)
                    if len(ts_str) == 19:
                        ts_str += ":000"
                    return ts_str
                return "9999-99-99 99:99:99:999"
            final_log.sort(key=extract_sort_key, reverse=True)
            with open(signal_file, 'w', encoding='utf-8') as f:
                f.write('\n'.join(final_log))
            send_file_via_base64(conn, signal_file)
            os.unlink(signal_file)
        driver.quit()
    except Exception as e:
        conn.sendall(f"[EXCEPTION] collect_signal_logs 失败: {e}\n".encode())

def collect_scpi_log(conn):
    try:
        driver = create_driver()
        driver.get(URL)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        msg_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Message']"))
        )
        driver.execute_script("arguments[0].click();", msg_btn)
        time.sleep(1)
        scpi_panel = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "dataScpiBox"))
        )
        scpi_text = scpi_panel.text
        if scpi_text:
            lines = scpi_text.splitlines()
            filtered = [l for l in lines if l.strip() not in ("Scpi Message", "Copy All", "Clear")]
            final_text = "\n".join(filtered)
            ts = get_timestamp()
            scpi_file = f"/tmp/{ts}_scpi_message_log.txt"
            with open(scpi_file, 'w', encoding='utf-8') as f:
                f.write(final_text)
            send_file_via_base64(conn, scpi_file)
            os.unlink(scpi_file)
        driver.quit()
    except Exception as e:
        conn.sendall(f"[EXCEPTION] collect_scpi_log 失败: {e}\n".encode())

def do_pvt_screenshot(conn):
    try:
        driver = create_driver()
        driver.get(URL)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        time.sleep(1)
        measurement_btn = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Measurement')]"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", measurement_btn)
        driver.execute_script("arguments[0].click();", measurement_btn)
        time.sleep(1)
        main_screen_btn = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'MainScreen:')]/following-sibling::div//button"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", main_screen_btn)
        driver.execute_script("arguments[0].click();", main_screen_btn)
        time.sleep(1)
        script_pvt = """
        function clickPVT() {
            const allElements = document.querySelectorAll('*');
            for (let el of allElements) {
                if (el.textContent && el.textContent.trim() === 'PVT') {
                    let target = el;
                    for (let i = 0; i < 5; i++) {
                        if (target.classList && target.classList.contains('el-dropdown-menu__item')) {
                            target.click();
                            return true;
                        }
                        if (target.parentElement) {
                            target = target.parentElement;
                        } else break;
                    }
                }
            }
            return false;
        }
        return clickPVT();
        """
        if not driver.execute_script(script_pvt):
            items = driver.find_elements(By.XPATH, "//li[contains(@class, 'el-dropdown-menu__item')]")
            if len(items) >= 3:
                driver.execute_script("arguments[0].click();", items[2])
        time.sleep(1)
        blank_area = driver.find_element(By.XPATH, "//div[contains(@class, 'articleCenter')]")
        driver.execute_script("arguments[0].click();", blank_area)
        time.sleep(2)
        ts = get_timestamp()
        screenshot_path = f"/tmp/{ts}_PVT.png"
        driver.save_screenshot(screenshot_path)
        send_file_via_base64(conn, screenshot_path)
        os.unlink(screenshot_path)
        driver.quit()
    except Exception as e:
        conn.sendall(f"[EXCEPTION] PVT 截图失败: {e}\n".encode())

def do_lte_screenshot(conn):
    try:
        driver = create_driver()
        driver.get(URL)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        time.sleep(1)
        measurement_btn = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Measurement')]"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", measurement_btn)
        driver.execute_script("arguments[0].click();", measurement_btn)
        time.sleep(1)
        main_screen_btn = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'MainScreen:')]/following-sibling::div//button"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", main_screen_btn)
        driver.execute_script("arguments[0].click();", main_screen_btn)
        time.sleep(1)
        script_lte = """
        function clickLTE() {
            const allElements = document.querySelectorAll('*');
            for (let el of allElements) {
                if (el.textContent && el.textContent.trim() === 'LTE') {
                    let target = el;
                    for (let i = 0; i < 5; i++) {
                        if (target.classList && target.classList.contains('el-dropdown-menu__item')) {
                            target.click();
                            return true;
                        }
                        if (target.parentElement) {
                            target = target.parentElement;
                        } else break;
                    }
                }
            }
            return false;
        }
        return clickLTE();
        """
        if not driver.execute_script(script_lte):
            items = driver.find_elements(By.XPATH, "//li[contains(@class, 'el-dropdown-menu__item')]")
            if len(items) >= 3:
                driver.execute_script("arguments[0].click();", items[2])
        time.sleep(1)
        blank_area = driver.find_element(By.XPATH, "//div[contains(@class, 'articleCenter')]")
        driver.execute_script("arguments[0].click();", blank_area)
        time.sleep(2)
        ts = get_timestamp()
        screenshot_path = f"/tmp/{ts}_LTE.png"
        driver.save_screenshot(screenshot_path)
        send_file_via_base64(conn, screenshot_path)
        os.unlink(screenshot_path)
        driver.quit()
    except Exception as e:
        conn.sendall(f"[EXCEPTION] LTE 截图失败: {e}\n".encode())

def collect_core_network_logs(conn, kwargs):
    log_dir = kwargs.get('path', '/opt/mt5gc/var/log/mt5gc')
    try:
        report_resource(conn, "收集核心网日志前")
        if not os.path.isdir(log_dir):
            conn.sendall(f"[ERROR] 核心网日志目录不存在: {log_dir}\n".encode())
            return
        try:
            file_count = len(os.listdir(log_dir))
            conn.sendall(f"找到 {file_count} 个核心网日志文件，正在打包...\n".encode())
        except Exception:
            pass

        ts = get_timestamp()
        archive_path = f"/tmp/{ts}_core_network_logs.tar"
        res = subprocess.run(
            f"echo '{SUDO_PASSWORD}' | sudo -S tar -cf {archive_path} -C {log_dir} .",
            shell=True, capture_output=True, text=True, timeout=120
        )
        if res.returncode != 0 or not os.path.exists(archive_path):
            conn.sendall(f"[ERROR] 打包核心网日志失败: {res.stderr.strip()}\n".encode())
            return

        size_mb = os.path.getsize(archive_path) / 1024 / 1024
        log_with_timestamp(f"核心网日志打包完成: {archive_path} ({size_mb:.2f} MB)")
        send_file_via_base64(conn, archive_path, compress_threshold_mb=1)

        try:
            subprocess.run(f"echo '{SUDO_PASSWORD}' | sudo -S rm -f {archive_path}",
                           shell=True, timeout=5)
        except Exception as e:
            log_with_timestamp(f"删除核心网日志压缩包失败: {e}")
        conn.sendall("[OK] 核心网日志发送完成\n".encode())
    except Exception as e:
        conn.sendall(f"[EXCEPTION] collect_core_network_logs 失败: {e}\n".encode())

def collect_current_logs(conn):
    log_dir = "/tmp/yc1100/current"
    try:
        if not os.path.isdir(log_dir):
            conn.sendall(f"[ERROR] 基站 current 日志目录不存在: {log_dir}\n".encode())
            return

        import glob
        all_files = os.listdir(log_dir)
        if not all_files:
            conn.sendall("[ERROR] 日志目录为空\n".encode())
            return

        selected = []
        mainctrl_path = os.path.join(log_dir, "mainctrl.log")
        if os.path.exists(mainctrl_path):
            selected.append("mainctrl.log")
        else:
            main_files = [f for f in all_files if "mainctrl" in f.lower()]
            if main_files:
                selected.append(main_files[0])

        nr_files = [f for f in all_files if "NR" in f and f.endswith(".log")]
        if nr_files:
            nr_files.sort(key=lambda f: os.path.getmtime(os.path.join(log_dir, f)), reverse=True)
            selected.append(nr_files[0])

        lte_files = [f for f in all_files if "LTE" in f and f.endswith(".log")]
        if lte_files:
            lte_files.sort(key=lambda f: os.path.getmtime(os.path.join(log_dir, f)), reverse=True)
            selected.append(lte_files[0])

        if not selected:
            conn.sendall("[ERROR] 未找到 mainctrl.log 或 NR/LTE 日志\n".encode())
            return

        conn.sendall(f"选中 {len(selected)} 个文件: {', '.join(selected)}\n".encode())

        ts = get_timestamp()
        archive_path = f"/tmp/{ts}_gnb_current_logs.tar"

        files_str = " ".join(selected)
        res = subprocess.run(
            f"echo '{SUDO_PASSWORD}' | sudo -S tar -cf {archive_path} -C {log_dir} {files_str}",
            shell=True, capture_output=True, text=True, timeout=60
        )

        if res.returncode != 0 or not os.path.exists(archive_path):
            conn.sendall(f"[ERROR] 打包 current 日志失败: {res.stderr.strip()}\n".encode())
            return

        size_mb = os.path.getsize(archive_path) / 1024 / 1024
        log_with_timestamp(f"基站 current 日志打包完成(mainctrl + 最新 NR/LTE): {archive_path} ({size_mb:.2f} MB)")
        send_file_via_base64(conn, archive_path, compress_threshold_mb=1)

        try:
            subprocess.run(f"echo '{SUDO_PASSWORD}' | sudo -S rm -f {archive_path}", shell=True, timeout=5)
        except Exception as e:
            log_with_timestamp(f"删除 current 日志压缩包失败: {e}")

        conn.sendall("[OK] 基站 current 日志发送完成\n".encode())

    except Exception as e:
        conn.sendall(f"[EXCEPTION] collect_current_logs 失败: {e}\n".encode())


def start_continuous_logging(conn):
    global _log_thread, _log_stop_event, _log_file_path
    global _scpi_driver, _scpi_ready, _scpi_lines

    if _log_thread and _log_thread.is_alive():
        conn.sendall("[WARN] 已有抓取线程在运行，请先 stop\n".encode())
        return

    if _scpi_driver:
        try:
            _scpi_driver.quit()
        except:
            pass
        _scpi_driver = None

    _log_stop_event = threading.Event()
    _log_file_path = None
    _scpi_lines = []
    _scpi_ready = True

    def capture_signal_and_scpi():
        global _log_file_path, _scpi_lines
        driver = None
        try:
            driver = create_driver()
            driver.get(URL)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            msg_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Message']"))
            )
            driver.execute_script("arguments[0].click();", msg_btn)
            time.sleep(0.5)

            try:
                scpi_panel = driver.find_element(By.ID, "dataScpiBox")
                clear_btn = scpi_panel.find_element(By.XPATH, ".//button[normalize-space()='Clear']")
                clear_btn.click()
                time.sleep(0.3)
            except Exception as e:
                log_with_timestamp(f"清空 SCPI 面板失败(不影响抓取): {e}")

            ts = get_timestamp()
            _log_file_path = f"/tmp/{ts}_signal_message_log.txt"
            seen_lines = set()
            final_log = []
            scpi_seen = set()

            while not _log_stop_event.is_set():
                current_text = driver.execute_script("return document.body.innerText;")
                for line in current_text.split('\n'):
                    line = line.strip()
                    if re.search(r'\d{4}-\d{2}-\d{2}', line) and 'Send->' not in line and 'Recv<-' not in line:
                        if line and line not in seen_lines:
                            seen_lines.add(line)
                            final_log.append(line)

                try:
                    scpi_text = driver.execute_script(
                        "var e=document.getElementById('dataScpiBox'); return e ? e.innerText : '';")
                    for line in scpi_text.splitlines():
                        s = line.strip()
                        if s and s not in ("Scpi Message", "Copy All", "Clear") and s not in scpi_seen:
                            scpi_seen.add(s)
                            _scpi_lines.append(s)
                except Exception:
                    pass

                time.sleep(0.5)

            timestamp_pattern = re.compile(r'(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}(?::\d{3})?)')
            def extract_sort_key(line):
                match = timestamp_pattern.search(line)
                if match:
                    ts_str = match.group(1)
                    if len(ts_str) == 19:
                        ts_str += ":000"
                    return ts_str
                return "9999-99-99 99:99:99:999"
            final_log.sort(key=extract_sort_key, reverse=True)
            filtered_log = [
                line for line in final_log
                if not re.match(r'^\d{4}-\d{2}-\d{2}(\s+\d{2}:\d{2}:\d{2}(:\d{3})?)?$', line.strip())
            ]
            with open(_log_file_path, 'w', encoding='utf-8') as f:
                f.write('\n'.join(filtered_log))
        except Exception as e:
            log_with_timestamp(f"Signal/SCPI 抓取线程异常: {e}")
        finally:
            if driver:
                driver.quit()

    _log_thread = threading.Thread(target=capture_signal_and_scpi, daemon=True)
    _log_thread.start()

    conn.sendall("[OK] 已启动持续抓取 Signal+SCPI 日志\n".encode())

def stop_continuous_logging(conn):
    global _log_thread, _log_stop_event, _log_file_path
    global _scpi_driver, _scpi_ready, _scpi_lines

    if _log_thread and _log_thread.is_alive():
        _log_stop_event.set()
        _log_thread.join(timeout=10)
        if _log_file_path and os.path.exists(_log_file_path):
            send_file_via_base64(conn, _log_file_path)
            os.unlink(_log_file_path)
    else:
        conn.sendall("[WARN] 没有正在运行的 Signal 抓取线程\n".encode())

    if _scpi_lines:
        ts = get_timestamp()
        scpi_file = f"/tmp/{ts}_scpi_message_log.txt"
        with open(scpi_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(_scpi_lines))
        send_file_via_base64(conn, scpi_file)
        os.unlink(scpi_file)
        log_with_timestamp(f"SCPI 日志共 {len(_scpi_lines)} 行")
        _scpi_lines = []
    else:
        log_with_timestamp("本轮未积累到 SCPI 行，尝试一次性补抓")
        try:
            conn.sendall("[WARN] SCPI 无累积内容，尝试一次性补抓\n".encode())
        except:
            pass
        collect_scpi_log(conn)

    if _scpi_driver:
        try:
            _scpi_driver.quit()
        except:
            pass
        _scpi_driver = None

def execute_shell_command(conn, cmd: str):
    try:
        proc = subprocess.Popen(
            cmd,
            shell=True,
            executable='/bin/bash',
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            stdin=subprocess.PIPE,
            start_new_session=True,
            bufsize=1,
            universal_newlines=True
        )
        proc.stdin.close()
        for line in proc.stdout:
            line = line.rstrip('\n')
            if line:
                stamped = log_with_timestamp(line)
                try:
                    conn.sendall((stamped + '\n').encode('utf-8'))
                except:
                    break
        proc.stdout.close()
        proc.wait()
    except Exception as e:
        conn.sendall(f"[EXCEPTION] 执行命令异常: {e}\n".encode())

def handle_diag_command(conn, raw_cmd: str):
    if '?' in raw_cmd[5:]:
        action_part, query_part = raw_cmd[5:].split('?', 1)
    else:
        action_part = raw_cmd[5:]
        query_part = ''
    action = action_part.lower()
    kwargs = {}
    if query_part:
        for param in query_part.split('&'):
            if '=' in param:
                k, v = param.split('=', 1)
                kwargs[k] = v
    if 'duration' not in kwargs:
        kwargs['duration'] = 60

    if action == 'restart':
        do_restart(conn)
    elif action == 'configure_log_and_restart':
        do_configure_log_and_restart(conn, kwargs)
    elif action == 'signal_logs':
        collect_signal_logs(conn, kwargs.get('duration', 60))
    elif action == 'scpi_log':
        collect_scpi_log(conn)
    elif action == 'pvt':
        do_pvt_screenshot(conn)
    elif action == 'lte':
        do_lte_screenshot(conn)
    elif action == 'start_log':
        start_continuous_logging(conn)
    elif action == 'stop_log':
        stop_continuous_logging(conn)
    elif action == 'collect_core_logs':
        collect_core_network_logs(conn, kwargs)
    elif action == 'collect_current_logs':
        collect_current_logs(conn)
    else:
        conn.sendall(f"未知的 DIAG action: {action}\n".encode())

def handle_client(conn, addr):
    try:
        command = b''
        while True:
            chunk = conn.recv(1)
            if not chunk:
                break
            if chunk == b'\n':
                break
            command += chunk
        cmd_str = command.decode('utf-8').strip()
        if not cmd_str:
            conn.sendall(b"Empty command\n")
            return
        if cmd_str.startswith("DIAG:"):
            handle_diag_command(conn, cmd_str)
        else:
            execute_shell_command(conn, cmd_str)
    except Exception as e:
        pass
    finally:
        conn.close()

def main():
    log_with_timestamp(f"===== ubuntu_server 启动 (VERSION={VERSION}) =====")
    server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    server.bind((HOST, PORT))
    server.listen(10)
    print(f"[*] 统一服务已启动，监听 {HOST}:{PORT}")
    while True:
        try:
            conn, addr = server.accept()
            threading.Thread(target=handle_client, args=(conn, addr), daemon=True).start()
        except KeyboardInterrupt:
            print("\n[!] 服务被用户中断")
            break
        except Exception as e:
            print(f"[-] 接受连接时出错: {e}")

if __name__ == "__main__":
    main()