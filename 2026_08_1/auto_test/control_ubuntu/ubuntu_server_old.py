import socket
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
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ========== 配置 ==========
HOST = '0.0.0.0'
PORT = 9999
FIREFOX_BINARY = "/snap/firefox/current/usr/lib/firefox/firefox"
GECKODRIVER_PATH = "/snap/bin/geckodriver"
URL = "http://127.0.0.1:8081/"
VERSION = "YC1100.1.00.02.19_alpha"
SUDO_PASSWORD = "yc"

# 服务器端持久运行日志：log_with_timestamp 的每一行都会追加到这里，
# 便于事后按 case 复盘（NR 退出/OOM/内存/文件收发等控制台信息不再随终端丢失）。
SERVER_LOG_FILE = "/tmp/ubuntu_server.log"
SERVER_LOG_MAX_BYTES = 50 * 1024 * 1024  # 超过则启动时轮转，避免无限增长

NR_READY_STRING = "Websocket Client Connected"
NR_STABILIZE_DELAY = 20

# ========== 全局变量 ==========
_log_thread = None
_log_stop_event = None
_log_file_path = None
_scpi_driver = None
_scpi_ready = False
_scpi_lines = []

_gnb_log_files = []
_gnb_log_handles = []
_gnb_log_handles_map = {}
_monitor_threads = []

# 当前测试 case 标识（由客户端 gNB 启动时经 DIAG 命令带入），用于服务器日志分段标注
_current_case = None
_server_log_lock = threading.Lock()
# 当前 case 在 SERVER_LOG_FILE 中的起始字节偏移(Master 启动时记下)，
# 会话结束(kill_processes)时据此把这一段服务器日志整段回传给客户端写入 souren_execution.log
_case_log_offset = 0

def log_with_timestamp(line: str):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    stamped = f"[{timestamp}] {line}"
    print(stamped, flush=True)
    # 同步落盘到持久日志文件，供事后按 case 复盘
    try:
        with _server_log_lock:
            with open(SERVER_LOG_FILE, 'a', encoding='utf-8') as f:
                f.write(stamped + '\n')
    except Exception:
        pass
    return stamped

def mark_case(kwargs, action: str):
    global _current_case
    raw = kwargs.get('case') if kwargs else None
    if raw:
        try:
            _current_case = base64.b64decode(raw).decode('utf-8')
        except Exception:
            _current_case = raw
    sep = "=" * 72
    log_with_timestamp(sep)
    log_with_timestamp(f"CASE [{_current_case or '未知'}] >>> {action}")
    log_with_timestamp(sep)

def send_case_server_log(conn):
    try:
        with _server_log_lock:
            with open(SERVER_LOG_FILE, 'r', encoding='utf-8', errors='ignore') as f:
                f.seek(_case_log_offset)
                data = f.read()
    except Exception as e:
        try:
            conn.sendall(f"!SLOG:[读取服务器日志失败] {e}\n".encode('utf-8'))
        except Exception:
            pass
        return
    for line in data.splitlines():
        try:
            conn.sendall(("!SLOG:" + line + "\n").encode('utf-8'))
        except Exception:
            break

def log_mem_status(conn, tag: str):
    try:
        mem = subprocess.run("free -h", shell=True, capture_output=True, text=True, timeout=5).stdout
    except Exception as e:
        mem = f"(获取内存失败: {e})"
    log_with_timestamp(f"内存状态[{tag}]:\n{mem}")
    if conn:
        try:
            conn.sendall(f"[MEM:{tag}] {mem}\n".encode())
        except Exception:
            pass

def classify_output_source(line: str) -> str:
    lte_patterns = [
        'lte-softmodem',
        'CMDLINE.*lte-softmodem',
        'lte_softmodem',
        'LTE Oai Version',
        'LTE.*started',
        'ALL RUs ready - ALL eNBs ready',
        '\[LTE_RRC\]',
        '\[LTE_MAC\]',
        '\[LTE_PHY\]',
        'Attach to eNB',
        'eNB L1.*configured',
        'eNB_APP',
        'LTE_READY',
        'LTE application initialized',
        'Initializing eNB',
    ]

    for pattern in lte_patterns:
        if re.search(pattern, line, re.IGNORECASE):
            return "lte"

    # NR patterns - 排除掉LTE后检查
    exclude_patterns = [
        'no process found',
        'lte-softmodem.*no process',
    ]
    for pattern in exclude_patterns:
        if re.search(pattern, line, re.IGNORECASE):
            return "nr"

    nr_patterns = [
        'nr-softmodem', 'CMDLINE.*nr-softmodem', 'NR Oai Version',
        'Entering ITTI signals handler', 'TYPE <CTRL-C> TO TERMINATE',
        'RF_Set_rf_enable', 'RF_SetTxLevel', 'RF_SetRxLevel',
        'OAI set_RF_Enable', 'nahaiy', 'set_channel_rf_enable',
        'nr_phy_config_request', 'DL frequency.*band',
        'got sync.*ru_thread', 'L1_stats_thread',
        '\[UTIL\].*threadCreate.*Tpool',
        '\[UTIL\].*threadCreate.*L1_',
        '\[UTIL\].*threadCreate.*TASK_',
        'ALL RUs ready.*gNBs ready',
        'START MAIN THREADS',
        'Initializing gNB threads',
        'wait_gNBs',
        'gNB L1.*configured',
        'RC\.nb_nr_L1_inst',
        'RC\.nb_RU',
        'About to Init RU threads',
        'Initializing RU threads',
        'init_eNB_afterRU',
        '\[LOADER\].*libdfts|libldpc',
        'shlib_path.*libdfts|libldpc',
        'Initialise nr transport',
        'Mapping RX ports.*RUs to gNB',
        'Attaching RU.*antenna.*gNB',
        'Sending sync to all threads',
        'waiting for sync.*L1_stats_thread',
        '\[MT_PS\]',
        '\[NR_RRC\]',
        '\[NR_MAC\]',
        '\[NR_PHY\]',
        '\[GNB_APP\]',
        '\[RRC\]',
        '\[MAC\]',
        '\[PHY\]',
        '\[ITTI\]',
        '\[NGAP\]',
        '\[GTPU\]',
        '\[SCTP\]',
        '\[X2AP\]',
        'ru_thread',
        'Exiting ru_thread',
        'set cell on',
        'ru->firstread_flag',
        'ps to mainctrl',
        'mainctrl to ps',
        'SendCellState_to_mainctrl',
        'Cell_Standalone_type',
        'mt_update.*config',
        'mt_mainctrl_cellon',
        'tdd_slot_info',
        'Update gNB_RU',
        'RCupdate_RU',
        '<BCCH-BCH-Message>',
        '<BCCH-DL-SCH-Message>',
        '<MeasurementTimingConfiguration>',
        'prach_I0',
        'UE.*RNTI',
        'RA-RNTI',
        'TC-RNTI',
        'Msg2|Msg3',
        'RAPROC',
        'Frame\.Slot',
        'SS TX:',
    ]

    for pattern in nr_patterns:
        if re.search(pattern, line, re.IGNORECASE):
            return "nr"
    return "master"

def monitor_file_and_add_timestamp(raw_file, final_file, stop_event, proc=None, ready_string=None, ready_event=None, nr_pid=None, conn=None, wrapper_proc=None):
    for _ in range(30):
        if os.path.exists(raw_file):
            break
        time.sleep(0.5)

    if not os.path.exists(raw_file):
        log_with_timestamp(f"原始日志文件未创建: {raw_file}")
        return

    exit_notified = False
    wrapper_exit_notified = False

    try:
        tail_proc = subprocess.Popen(
            ['tail', '-F', '-n', '+1', raw_file],
            stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL,
            bufsize=1,
            universal_newlines=True
        )

        with open(final_file, 'w', encoding='utf-8', buffering=1) as f_final:
            while True:
                if nr_pid and not exit_notified:
                    try:
                        os.kill(nr_pid, 0)
                    except PermissionError:
                        # nr-softmodem 以 root 身份运行（sudo），本进程是普通用户，
                        # 内核会先确认目标进程存在、再检查信号权限——EPERM 恰恰说明
                        # 它还活着，只是我们没权限对它做信号探测，绝不能当成"已退出"。
                        # 之前这里用笼统的 except OSError 把 EPERM 和 ESRCH 混为一谈，
                        # 导致 nr-softmodem 刚起来就被误判"已退出"，连带 OOM 诊断也在
                        # 错误的时间点触发，抓不到进程真正终止那一刻的系统日志。
                        pass
                    except ProcessLookupError:
                        exit_notified = True
                        exit_code = "未知"
                        try:
                            _, status = os.waitpid(nr_pid, os.WNOHANG)
                            if os.WIFSIGNALED(status):
                                exit_code = f"signal {os.WTERMSIG(status)}"
                            elif os.WIFEXITED(status):
                                exit_code = f"exit {os.WEXITSTATUS(status)}"
                        except:
                            pass
                        msg = f"NR 进程 (PID {nr_pid}) 退出，原因: {exit_code}"
                        log_with_timestamp(msg)
                        if conn:
                            try:
                                conn.sendall(f"[WARN] {msg}\n".encode())
                            except:
                                pass

                        try:
                            oom_proc = subprocess.run(
                                ["journalctl", "-k", "--no-pager", "-n", "100"],
                                capture_output=True, text=True, timeout=10
                            )
                            oom_lines = oom_proc.stdout.splitlines()
                            oom_filtered = [l for l in oom_lines if re.search(r'oom|out of memory|killed process', l, re.IGNORECASE)]
                            if oom_filtered:
                                for line in oom_filtered[-5:]:
                                    log_with_timestamp(f"OOM: {line}")
                                if conn:
                                    conn.sendall(f"[OOM] {chr(10).join(oom_filtered)}\n".encode())
                            else:
                                log_with_timestamp("未发现 OOM 日志（journalctl）")
                        except Exception as e:
                            log_with_timestamp(f"OOM 日志收集失败: {e}")

                        try:
                            mem_info = subprocess.run(["free", "-h"], capture_output=True, text=True, timeout=5)
                            log_with_timestamp(f"退出后内存状态:\n{mem_info.stdout}")
                            if conn:
                                conn.sendall(f"[MEM_AFTER] {mem_info.stdout}\n".encode())
                        except:
                            pass

                if wrapper_proc and not wrapper_exit_notified:
                    poll = wrapper_proc.poll()
                    if poll is not None:
                        wrapper_exit_notified = True
                        msg = f"包装脚本进程 (PID {wrapper_proc.pid}) 意外退出，退出码: {poll}"
                        log_with_timestamp(msg)
                        if conn:
                            try:
                                conn.sendall(f"[WARN] {msg}\n".encode())
                            except:
                                pass

                if stop_event.is_set():
                    try:
                        tail_proc.terminate()
                    except:
                        pass
                    while True:
                        line = tail_proc.stdout.readline()
                        if not line:
                            break
                        decoded_line = line.rstrip('\n\r')
                        if decoded_line:
                            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                            f_final.write(f"[{timestamp}]{decoded_line}\n")
                            f_final.flush()
                            if ready_string and ready_event and not ready_event.is_set():
                                if ready_string in decoded_line:
                                    ready_event.set()
                    break

                rlist, _, _ = select.select([tail_proc.stdout], [], [], 0.5)
                if rlist:
                    line = tail_proc.stdout.readline()
                    if not line:
                        break
                    decoded_line = line.rstrip('\n\r')
                    if decoded_line:
                        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        f_final.write(f"[{timestamp}]{decoded_line}\n")
                        f_final.flush()
                        if ready_string and ready_event and not ready_event.is_set():
                            if ready_string in decoded_line:
                                ready_event.set()

        try:
            tail_proc.terminate()
            tail_proc.wait(timeout=2)
        except:
            tail_proc.kill()

    except Exception as e:
        log_with_timestamp(f"监控文件异常 {raw_file}: {e}")

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

def do_restart(conn):
    try:
        driver = create_driver()
        driver.get(URL)
        wait = WebDriverWait(driver, 20)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        time.sleep(1)
        restart_xpath = "//div[text()='Restart' and contains(@class, 'headBtn')]"
        btn = driver.find_element(By.XPATH, restart_xpath)
        btn.click()
        try:
            WebDriverWait(driver, 2).until(EC.alert_is_present())
            alert = driver.switch_to.alert
            alert.accept()
        except:
            pass
        time.sleep(0.5)
        possible_ok = [
            "//button[contains(@class, 'el-button--primary') and contains(translate(., 'OK', 'ok'), 'ok')]",
            "//button[contains(translate(., 'OK', 'ok'), 'ok')]",
            "//div[contains(@class, 'el-message-box')]//button[contains(@class, 'el-button--primary')]",
            "//button[normalize-space(text())='OK']"
        ]
        for xp in possible_ok:
            try:
                btn = wait.until(EC.element_to_be_clickable((By.XPATH, xp)))
                btn.click()
                break
            except:
                continue
        driver.quit()
        conn.sendall("[OK] 重启完成\n".encode())
    except Exception as e:
        conn.sendall(f"[EXCEPTION] do_restart 失败: {e}\n".encode())

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

# ========== Master 启动 ==========
def execute_master_start(conn, kwargs=None):
    global _gnb_log_files, _gnb_log_handles, _monitor_threads, _case_log_offset
    try:
        _case_log_offset = os.path.getsize(SERVER_LOG_FILE) if os.path.exists(SERVER_LOG_FILE) else 0
    except Exception:
        _case_log_offset = 0
    mark_case(kwargs, "Master Control 启动")
    ts = get_timestamp()
    raw_log_file = f"/tmp/{ts}_master_raw.txt"
    final_log_file = f"/tmp/{ts}_master_control_log.txt"
    _gnb_log_files.append({"type": "master", "path": final_log_file})
    log_with_timestamp(f"Master 日志将写入: {final_log_file}")

    lib_path = f"./:/opt/yc1100/relversion/{VERSION}/release.dir/bin/Measurement/"
    inner_cmd = (f"cd /opt/yc1100/relversion/{VERSION}; "
                 f"export LD_LIBRARY_PATH={lib_path}; "
                 f"stdbuf -o0 -e0 sudo -S -E ./run.sh -o 1 2>&1")
    cmd = f"script -f -q {raw_log_file} -c '{inner_cmd}'"

    proc = subprocess.Popen(
        cmd,
        shell=True,
        executable='/bin/bash',
        stdin=subprocess.PIPE,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        start_new_session=True
    )
    try:
        proc.stdin.write((SUDO_PASSWORD + '\n').encode())
        proc.stdin.flush()
        proc.stdin.close()
    except:
        pass

    stop_event = threading.Event()
    ready_event = threading.Event()
    monitor_thread = threading.Thread(
        target=monitor_file_and_add_timestamp,
        args=(raw_log_file, final_log_file, stop_event, None, NR_READY_STRING, ready_event),
        daemon=True
    )
    monitor_thread.start()
    _monitor_threads.append({
        "thread": monitor_thread,
        "stop_event": stop_event,
        "proc": proc,
        "session_id": proc.pid,
        "final_file": final_log_file,
        "raw_file": raw_log_file
    })

    if ready_event.wait(timeout=60):
        log_with_timestamp(f"检测到就绪标志: {NR_READY_STRING}")
        log_mem_status(conn, "Master就绪")
        conn.sendall("[READY] Master Control started\n".encode())
    else:
        log_with_timestamp(f"超时未检测到就绪标志: {NR_READY_STRING}")
        conn.sendall("[WARN] Master Control not ready\n".encode())

# ========== RAN 启动（最终方案：包装脚本等待 NR 进程） ==========
def execute_ran_start(conn, kwargs):
    global _gnb_log_files, _gnb_log_handles, _monitor_threads
    mode = kwargs.get('mode', 'SA').upper()
    part = kwargs.get('part', 'NR')
    extra = kwargs.get('extra', '')
    mark_case(kwargs, f"RAN 启动 (mode={mode}, part={part})")
    # 默认改为 1（open）：startRan.sh 在 -o 0（close）时会让 nr-softmodem 脱离
    # 当前终端（daemonize），其真实输出完全不会流向我们外层 script -f -q 捕获的
    # 伪终端，这才是日志一直不完整的根本原因。-o 1 应该让 nr-softmodem 保持挂在
    # 启动它的终端上，从而被我们的 pty 捕获到。
    open_terminal = kwargs.get('ot', '1')

    if mode == 'SA':
        base_name = "sa_nr"
        cmd_args = f"-m SA -p NR -o {open_terminal}"
    elif mode == 'NSA':
        if part.upper() == 'NR':
            base_name = "nsa_nr"
            cmd_args = f"-m NSA -p NR -o {open_terminal}"
            if extra:
                cmd_args += f" --thread-pool \\\"{extra}\\\""
        elif part.upper() == 'LTE':
            base_name = "nsa_lte"
            cmd_args = f"-m NSA -p LTE -o {open_terminal}"
        else:
            conn.sendall(f"[ERROR] NSA模式需指定 part=NR 或 part=LTE\n".encode())
            return
    else:
        conn.sendall(f"[ERROR] 仅支持 SA 和 NSA 模式\n".encode())
        return

    base_dir = f"/opt/yc1100/relversion/{VERSION}"
    script_dir = base_dir
    startran_script = f"{script_dir}/startRan.sh"
    if not os.path.exists(startran_script):
        conn.sendall(f"[ERROR] startRan.sh 不存在: {startran_script}\n".encode())
        return

    ts = get_timestamp()
    raw_log_file = f"/tmp/{ts}_{base_name}_raw.txt"
    final_log_file = f"/tmp/{ts}_{base_name}_log.txt"
    log_type = "lte" if base_name == "nsa_lte" else "nr"
    _gnb_log_files.append({"type": log_type, "path": final_log_file})
    log_with_timestamp(f"{base_name.upper()} 日志将写入: {final_log_file}")

    # 选择进程检测模式（NR 或 LTE）
    if base_name in ["sa_nr", "nsa_nr"]:
        process_pattern = "nr-softmodem"
        process_name = "nr-softmodem"
        exit_msg = "nr-softmodem 进程已全部退出"
    else:  # nsa_lte
        process_pattern = "lte-softmodem"
        process_name = "lte-softmodem"
        exit_msg = "lte-softmodem 进程已全部退出"

    # 它的存在是为了解决一个 Linux 系统编程中的经典难题：如何让父进程（Python）精准感知到孙进程（nr-softmodem）的死亡，并保证死亡瞬间的日志不丢失。
    #-1000：这是 OOM（Out Of Memory）分数调整的最小值。当系统内存耗尽时，Linux 内核会优先杀死分数高的进程。设置为 -1000 意味着除非系统彻底崩溃，
    # 否则内核永远不会杀掉这个 Wrapper 进程。这保证了看守者永远比 nr-softmodem 活得更久，避免看守者被杀后，外层 script 提前关闭管道导致日志截断。
    #PID 落盘：将当前 Bash 进程的 PID 写入 /tmp，供 Python 侧读取，用于后续的会话清理（pkill -9 -s {sid}）。
    wrapper_content = f"""#!/bin/bash 
echo -1000 > /proc/self/oom_score_adj
echo $$ > /tmp/wrapper_{ts}.pid

echo "--- 启动 startRan.sh ({process_name}) ---"
cd "{script_dir}"
stdbuf -o0 -e0 ./startRan.sh {cmd_args} &
START_PID=$!

# 等待目标进程出现（最长 15 秒）
TARGET_PID=""
for i in $(seq 1 15); do
    TARGET_PID=$(pgrep -f {process_pattern} | head -1)
    if [ -n "$TARGET_PID" ]; then
        break
    fi
    sleep 1
done

if [ -z "$TARGET_PID" ]; then
    echo "ERROR: 无法检测到 {process_pattern} 进程"
    wait $START_PID
    exit 1
fi

echo "{process_name} PID: $TARGET_PID"
# 持续轮询系统中是否还存在任意目标进程，不依赖某个具体 PID
while pgrep -f {process_pattern} >/dev/null 2>&1; do
    sleep 1
done
echo "--- {exit_msg} ---"
# 确保 startRan.sh 也被清理
kill $START_PID 2>/dev/null
wait $START_PID 2>/dev/null
sleep infinity
"""

    with tempfile.NamedTemporaryFile(mode='w', suffix='.sh', delete=False, dir='/tmp') as f:
        f.write(wrapper_content)
        wrapper_path = f.name
    os.chmod(wrapper_path, 0o755)
    log_with_timestamp(f"包装脚本已创建: {wrapper_path}")

    # script -f -q 分配伪终端捕获整个包装脚本会话的输出（含 startRan.sh/nr-softmodem），
    # -f 持续刷新，避免 nr-softmodem 异常退出过快时日志被截断
    cmd = f"sudo -S script -f -q {raw_log_file} -c 'bash {wrapper_path}'"
    proc = subprocess.Popen(
        cmd,
        shell=True,
        executable='/bin/bash',
        stdin=subprocess.PIPE,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        start_new_session=True
    )
    try:
        proc.stdin.write((SUDO_PASSWORD + '\n').encode())
        proc.stdin.flush()
        proc.stdin.close()
    except:
        pass

    # 读取包装脚本 PID
    wrapper_pid = None
    pid_file = f"/tmp/wrapper_{ts}.pid"
    for _ in range(5):
        time.sleep(0.5)
        if os.path.exists(pid_file):
            try:
                with open(pid_file, 'r') as pf:
                    wrapper_pid = int(pf.read().strip())
                log_with_timestamp(f"包装脚本 PID: {wrapper_pid}")
                break
            except:
                pass

    # 等待目标进程出现（Python 侧快速确认）
    target_pid = None
    for _ in range(15):
        time.sleep(1)
        try:
            res = subprocess.run(["pgrep", "-f", process_pattern], capture_output=True, text=True)
            pids = res.stdout.strip().split('\n')
            if pids and pids[0]:
                target_pid = int(pids[0])
                log_with_timestamp(f"检测到 {process_pattern} PID: {target_pid}")
                break
        except:
            pass

    if not target_pid:
        log_with_timestamp(f"未能检测到 {process_pattern} 进程 PID")
        conn.sendall(f"[WARN] {process_name.upper()} 进程可能未正常启动\n".encode())

    stop_event = threading.Event()
    monitor_thread = threading.Thread(
        target=monitor_file_and_add_timestamp,
        args=(raw_log_file, final_log_file, stop_event, None, None, None, target_pid, conn, proc),
        daemon=True
    )
    monitor_thread.start()
    _monitor_threads.append({
        "thread": monitor_thread,
        "stop_event": stop_event,
        "proc": proc,
        "session_id": proc.pid,
        "final_file": final_log_file,
        "raw_file": raw_log_file,
        "wrapper_path": wrapper_path,
        "pid_file": pid_file
    })

    if mode == 'SA':
        log_with_timestamp("等待 NR 启动（最长 15 秒）...")
        nr_up = False
        for _ in range(15):
            time.sleep(1)
            if subprocess.run("pgrep -f 'nr-softmodem'", shell=True, capture_output=True).returncode == 0:
                nr_up = True
                break
        if nr_up:
            # 进程存在 ≠ 小区就绪。手动启动"总能连上"是因为人会等终端把初始化
            # 输出刷完才开测；这里改为继续轮询日志中的真实就绪标志，等到了才回
            # READY，避免框架在小区未就绪时就往下走导致 UE 连不上。
            READY_MARKERS = ("ALL RUs ready", "got sync", "Sending sync to all threads")
            marker_found = None
            for _ in range(60):
                try:
                    with open(raw_log_file, 'r', errors='ignore') as rf:
                        content = rf.read()
                    for m in READY_MARKERS:
                        if m in content:
                            marker_found = m
                            break
                except Exception:
                    pass
                if marker_found:
                    break
                # 进程中途挂掉则立即失败，不再干等
                if subprocess.run("pgrep -f 'nr-softmodem'", shell=True, capture_output=True).returncode != 0:
                    break
                time.sleep(1)

            if marker_found:
                log_with_timestamp(f"检测到 NR 就绪标志: {marker_found}")
                conn.sendall("[READY] SA NR started\n".encode())
                log_mem_status(conn, "SA-NR启动后")
                time.sleep(NR_STABILIZE_DELAY)
            elif subprocess.run("pgrep -f 'nr-softmodem'", shell=True, capture_output=True).returncode == 0:
                log_with_timestamp("60 秒未见就绪标志，但 NR 进程仍在，按就绪返回（客户端将延长稳定等待）")
                conn.sendall("[READY] SA NR started\n".encode())
                log_mem_status(conn, "SA-NR启动后")
                time.sleep(NR_STABILIZE_DELAY)
            else:
                log_with_timestamp("NR 进程已退出，启动失败")
                conn.sendall("[ERROR] SA NR start failed\n".encode())
        else:
            conn.sendall("[ERROR] SA NR start failed\n".encode())
    elif mode == 'NSA' and part.upper() == 'NR':
        log_with_timestamp("等待 NSA NR 启动（最长 15 秒）...")
        nr_up = False
        for _ in range(15):
            time.sleep(1)
            if subprocess.run("pgrep -f 'nr-softmodem'", shell=True, capture_output=True).returncode == 0:
                nr_up = True
                break
        if nr_up:
            conn.sendall("[READY] NSA NR started\n".encode())
            log_mem_status(conn, "NSA-NR启动后")
        else:
            conn.sendall("[ERROR] NSA NR start failed\n".encode())
    elif mode == 'NSA' and part.upper() == 'LTE':
        log_with_timestamp("等待 NSA LTE 启动（最长 15 秒）...")
        lte_up = False
        for _ in range(15):
            time.sleep(1)
            if subprocess.run("pgrep -f 'lte-softmodem'", shell=True, capture_output=True).returncode == 0:
                lte_up = True
                break
        if lte_up:
            conn.sendall("[READY] NSA LTE started\n".encode())
            log_mem_status(conn, "NSA-LTE启动后")
        else:
            conn.sendall("[ERROR] NSA LTE start failed\n".encode())

# ========== 核心网日志收集 ==========
def collect_core_network_logs(conn, kwargs):
    # 把 /opt/mt5gc/var/log/mt5gc 下全部日志用 tar 一次性打包后单文件回传。
    # 之前逐个文件 base64 传输 16 个 log 太慢（每个文件一次连接/编码往返）；
    # 打包一次、传输一次即可。压缩交给 send_file_via_base64 的 gzip 逻辑
    # （阈值降到 1MB，文本日志压缩率高），客户端按既有 .gz 逻辑自动解压出 .tar。
    log_dir = kwargs.get('path', '/opt/mt5gc/var/log/mt5gc')
    try:
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
        # 日志可能属 root，用 sudo 打包
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

def execute_cleanup_wrappers(conn):
    global _gnb_log_files, _gnb_log_handles, _gnb_log_handles_map, _monitor_threads
    log_with_timestamp("中断清理：回传日志、停 tail 读取器，保留基站及其 script/wrapper 父进程")

    for info in _monitor_threads:
        try:
            info["stop_event"].set()
        except Exception:
            pass
    for info in _monitor_threads:
        try:
            info["thread"].join(timeout=10)
        except Exception:
            pass

    time.sleep(1)

    for log_info in _gnb_log_files:
        send_file_via_base64(conn, log_info["path"])
        try:
            os.unlink(log_info["path"])
        except Exception:
            pass
        raw_path = log_info["path"].replace("_log.txt", "_raw.txt").replace("control_log", "raw")
        if os.path.exists(raw_path):
            try:
                subprocess.run(f"echo '{SUDO_PASSWORD}' | sudo -S -p '' rm -f {raw_path}", shell=True, timeout=5)
            except Exception:
                pass
    for info in _monitor_threads:
        for k in ("wrapper_path", "pid_file"):
            p = info.get(k)
            if p and os.path.exists(p):
                try:
                    os.unlink(p)
                except Exception:
                    pass

    _gnb_log_files.clear()
    _gnb_log_handles.clear()
    _gnb_log_handles_map.clear()
    _monitor_threads.clear()

    kill_cmd = (
        "pkill -9 -f 'script -f -q /tmp/.*_raw.txt' 2>/dev/null || true; "
        "pkill -9 -f 'tail -F.*/tmp/.*_raw.txt' 2>/dev/null || true; "
        "pkill -9 -f 'bash /tmp/tmp.*[.]sh' 2>/dev/null || true"
    )
    full_cmd = f"echo '{SUDO_PASSWORD}' | sudo -S -p '' bash -c \"{kill_cmd}\""
    try:
        subprocess.run(full_cmd, shell=True, executable='/bin/bash', capture_output=True, timeout=10)
    except Exception as e:
        log_with_timestamp(f"清理 plumbing 异常: {e}")

    # 4) 回传本 case 的服务器操作日志(与正常 kill 一致)
    send_case_server_log(conn)

    log_with_timestamp("中断清理完成(日志已回传，plumbing 已清，基站 setsid 独立存活)")
    try:
        conn.sendall("[OK] 日志已回传，plumbing 已清，基站保留继续运行\n".encode())
    except Exception:
        pass

def execute_kill_processes(conn):
    global _gnb_log_files, _gnb_log_handles, _gnb_log_handles_map, _monitor_threads
    log_with_timestamp(f"准备停止所有监控线程并发送 {len(_gnb_log_files)} 个日志文件")

    # 第一步：先礼后兵。nr-softmodem/lte-softmodem 自己支持 SIGINT 优雅退出
    # （它启动时会打印 "TYPE <CTRL-C> TO TERMINATE"），这里先单独发 SIGINT 并
    # 等它真正退出，这样 wrapper 脚本里的 pgrep 轮询才有机会打印
    # "--- nr-softmodem 进程已全部退出 ---" 之类的收尾信息，且这几秒的输出还能
    # 被下面仍在运行的 tail -F 采集到。绝不能在这一步之前就 set stop_event 或
    # 做会话级 SIGKILL，否则日志会在进程真正退出前就被截断（此前的 bug）。
    try:
        subprocess.run(
            f"echo '{SUDO_PASSWORD}' | sudo -S pkill -INT -f 'nr-softmodem|lte-softmodem'",
            shell=True, capture_output=True, timeout=5
        )
    except Exception as e:
        log_with_timestamp(f"优雅停止 softmodem 失败: {e}")

    for _ in range(16):
        res = subprocess.run("pgrep -f 'nr-softmodem|lte-softmodem'", shell=True, capture_output=True)
        if res.returncode != 0:
            break
        time.sleep(0.5)

    # 再等一小会，确保 script -f 已把收尾输出 flush 到 raw 文件、tail -F 也已读到，
    # 避免紧接着的会话级清理把还没被 tail 读走的尾部内容截断
    time.sleep(1)

    for info in _monitor_threads:
        proc = info.get("proc")
        if proc and proc.poll() is None:
            try:
                os.kill(proc.pid, signal.SIGINT)
            except:
                pass
            time.sleep(0.5)
            if proc.poll() is None:
                try:
                    proc.terminate()
                except:
                    pass
            try:
                proc.wait(timeout=3)
            except:
                proc.kill()
        wrapper = info.get("wrapper_path")
        if wrapper and os.path.exists(wrapper):
            try:
                os.unlink(wrapper)
                log_with_timestamp(f"已删除包装脚本: {wrapper}")
            except:
                pass
        pid_file = info.get("pid_file")
        if pid_file and os.path.exists(pid_file):
            try:
                os.unlink(pid_file)
            except:
                pass

    # 按会话 ID 做一次确定性的提权清理：proc 启动时使用了 start_new_session=True，
    # 其会话内所有子孙进程（sudo/script/tee/wrapper.sh/startRan.sh/nr-softmodem 等，
    # 即便以 root 身份运行）都共享同一个会话 ID，无需依赖信号逐级转发即可一次性杀干净。
    # 此时 nr-softmodem 理应已经通过上面的 SIGINT 优雅退出，这里的 SIGKILL 主要是
    # 清理 wrapper.sh（卡在 sleep infinity）/script/tee/sudo 等残留会话进程。
    for info in _monitor_threads:
        sid = info.get("session_id")
        if sid:
            try:
                subprocess.run(
                    f"echo '{SUDO_PASSWORD}' | sudo -S pkill -9 -s {sid}",
                    shell=True, capture_output=True, timeout=5
                )
            except Exception as e:
                log_with_timestamp(f"按会话清理失败 (sid={sid}): {e}")

    # 会话已清理完毕，现在才通知监控线程停止 tail —— 保证进程收尾阶段的最后输出
    # 已经被读入 final 文件，不会像之前那样在进程真正退出前就被提前截断。
    for info in _monitor_threads:
        info["stop_event"].set()

    for info in _monitor_threads:
        try:
            info["thread"].join(timeout=10)
            if info["thread"].is_alive():
                log_with_timestamp(f"线程 {info['final_file']} 未在 10 秒内结束，强制继续")
        except Exception as e:
            log_with_timestamp(f"等待线程异常: {e}")

    time.sleep(1)

    log_mem_status(conn, "回传文件前")
    for log_info in _gnb_log_files:
        send_file_via_base64(conn, log_info["path"])
        try:
            os.unlink(log_info["path"])
            log_with_timestamp(f"已删除临时文件: {log_info['path']}")
        except Exception as e:
            log_with_timestamp(f"删除文件失败 {log_info['path']}: {e}")
        raw_path = log_info["path"].replace("_log.txt", "_raw.txt").replace("control_log", "raw")
        if os.path.exists(raw_path):
            try:
                # raw 文件由 root 创建，使用 sudo 删除
                subprocess.run(f"echo '{SUDO_PASSWORD}' | sudo -S rm -f {raw_path}", shell=True, timeout=5)
                log_with_timestamp(f"已删除原始文件: {raw_path}")
            except Exception as e:
                log_with_timestamp(f"删除原始文件失败 {raw_path}: {e}")

    _gnb_log_files.clear()
    _gnb_log_handles.clear()
    _gnb_log_handles_map.clear()
    _monitor_threads.clear()

    kill_cmd = (
        "pkill -9 -f 'nr-softmodem|lte-softmodem' 2>/dev/null || true; "
        # 现在的日志捕获用的是 `sudo -S script -f -q /tmp/*_raw.txt -c 'bash /tmp/tmpXXXX.sh'`，
        # 而多数发行版 sudo 默认 use_pty，会把 script 子进程放进独立会话，前面按
        # session_id 的 pkill -s 抓不到它，于是 sudo/script/wrapper 三件套会残留。
        # 这里按实际 cmdline 特征精确补杀 script 与 wrapper 临时脚本。
        "pkill -9 -f 'script -f -q /tmp/.*_raw.txt' 2>/dev/null || true; "
        "pkill -9 -f 'bash /tmp/tmp.*[.]sh' 2>/dev/null || true"
    )
    # 关键：外层用双引号包裹 bash -c，内层模式用单引号。此前外层也是单引号，和内层
    # 的 'nr-softmodem|lte-softmodem' 单引号发生嵌套冲突，命令被 shell 撕裂，才会打出
    # "lte-softmodem: command not found" / "script: cannot open /tmp/: Is a directory"
    # 这类无用报错。同时去掉了会自匹配、易出错的 `ps|grep|awk '{print $2}'` 兜底
    # （上面两条 pkill 已足够），并用 -p '' 抑制 [sudo] password 提示，保持日志干净。
    full_cmd = f"echo '{SUDO_PASSWORD}' | sudo -S -p '' bash -c \"{kill_cmd}\""
    try:
        proc = subprocess.Popen(
            full_cmd,
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
                log_with_timestamp(line)
        proc.stdout.close()
        proc.wait()
    except Exception as e:
        log_with_timestamp(f"kill 命令异常: {e}")

    # 把本 case 从 Master 启动以来的完整服务器操作日志回传给客户端，
    # 让 souren_execution.log 拿到和服务器控制台一致的详细信息
    send_case_server_log(conn)

    conn.sendall("[OK] 进程清理完成\n".encode())

# ========== 持续抓取（完整实现） ==========
# 单浏览器实例同时抓 Signal + SCPI：
# SCPI 面板内容靠页面 websocket 实时推送积累，只有"挂着面板"的那个浏览器
# 实例能收到。之前用第二个 Firefox 专抓 SCPI，一旦它起不来/收不到推送，
# SCPI 日志就整个丢失(且旧代码为空时连文件都不回传)。现在合并到抓 Signal
# 的同一个页面里，轮询时顺带累积 SCPI 行，stop 时两个文件一起回传。
def start_continuous_logging(conn):
    global _log_thread, _log_stop_event, _log_file_path
    global _scpi_driver, _scpi_ready, _scpi_lines

    if _log_thread and _log_thread.is_alive():
        conn.sendall("[WARN] 已有抓取线程在运行，请先 stop\n".encode())
        return

    # 清理上一轮残留的 SCPI 驱动（旧版遗留，防 Firefox 实例泄漏）
    if _scpi_driver:
        try:
            _scpi_driver.quit()
        except:
            pass
        _scpi_driver = None

    _log_stop_event = threading.Event()
    _log_file_path = None
    _scpi_lines = []
    _scpi_ready = True  # 不再有独立的 SCPI 驱动需要等待

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

            # 清空 SCPI 面板，确保只抓本轮内容
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
                # ---- Signal：全页文本按日期行过滤 ----
                current_text = driver.execute_script("return document.body.innerText;")
                for line in current_text.split('\n'):
                    line = line.strip()
                    if re.search(r'\d{4}-\d{2}-\d{2}', line) and 'Send->' not in line and 'Recv<-' not in line:
                        if line and line not in seen_lines:
                            seen_lines.add(line)
                            final_log.append(line)

                # ---- SCPI：直接读面板文本，按行去重累积(保持面板顺序) ----
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
    elif action == 'master_start':
        execute_master_start(conn, kwargs)
    elif action == 'ran_start':
        execute_ran_start(conn, kwargs)
    elif action == 'kill_processes':
        execute_kill_processes(conn)
    elif action == 'cleanup_wrappers':
        execute_cleanup_wrappers(conn)
    elif action == 'collect_core_logs':
        collect_core_network_logs(conn, kwargs)
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
    try:
        if os.path.exists(SERVER_LOG_FILE) and os.path.getsize(SERVER_LOG_FILE) > SERVER_LOG_MAX_BYTES:
            os.replace(SERVER_LOG_FILE, SERVER_LOG_FILE + ".old")
    except Exception:
        pass
    log_with_timestamp(f"===== ubuntu_server 启动 (VERSION={VERSION}) 持久日志: {SERVER_LOG_FILE} =====")
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