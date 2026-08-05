import socket
import base64
import codecs
import gzip
import os
import time
import tarfile
import logging
import json
import souren_config

class RemoteClient:
    def __init__(self, host=None, port=None, password=None):
        if host is None:
            host = souren_config.DEFAULT_IP
        if port is None:
            port = souren_config.REMOTE_SERVER_PORT
        self.host = host
        self.port = port
        self.password = password or souren_config.REMOTE_SUDO_PASSWORD
        self.case = None  # 当前测试 case 标识，由 start_gnb 设置，用于日志分段标注

    def _remote_log(self, msg, echo=True):
        """把远程基站/服务器的信息写入 souren_execution.log（经 SourenToolSet 日志器
        传播到根日志的文件处理器），并带上 case 便于事后区分是哪个用例。
        之前 RemoteClient 全用 print，这些信息只到控制台、进不了 souren_execution.log。"""
        if echo:
            print(msg)
        try:
            case = self.case or os.path.basename(os.path.normpath(os.getcwd()))
        except Exception:
            case = None
        text = f"[远程][{case}] {msg}" if case else f"[远程] {msg}"
        try:
            logger = logging.getLogger('SourenToolSet')
            up = str(msg).upper()
            if any(k in up for k in ('WARN', 'ERROR', 'OOM', 'EXCEPTION')) or '退出' in msg or '失败' in msg:
                logger.warning(text)
            else:
                logger.info(text)
        except Exception:
            pass

    def _send_command_and_check_ok(self, command, ok_string="[OK]"):
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(5)
            sock.connect((self.host, self.port))
            sock.sendall((command + '\n').encode('utf-8'))
            buffer = ""
            # 增量解码器：sock.recv 可能在多字节(中文)字符中间切断，逐片 data.decode
            # 会崩("unexpected end of data")。用它把不完整的尾字节留到下一片再拼。
            _decoder = codecs.getincrementaldecoder('utf-8')()
            success = False
            while True:
                try:
                    data = sock.recv(8192)
                    if not data:
                        break
                    decoded = _decoder.decode(data)
                    buffer += decoded
                    if ok_string in buffer:
                        success = True
                    # 转发服务端整行输出到 souren_execution.log(经 _remote_log 按 case 名归档)，
                    # 使 do_restart 里的 [MEM] 内存快照等信息不再只停留在服务器控制台。
                    while '\n' in buffer:
                        line, buffer = buffer.split('\n', 1)
                        line = line.rstrip('\r')
                        if line.strip():
                            self._remote_log(line, echo=False)
                except socket.timeout:
                    continue
                except (ConnectionResetError, BrokenPipeError):
                    break
            if buffer.strip():
                self._remote_log(buffer.strip(), echo=False)
            sock.close()
            return success
        except Exception as e:
            print(f"[!] 发送命令失败: {e}")
            return False

    def _receive_file_from_command(self, command, timeout=60):
        sock = None
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(timeout)
            sock.connect((self.host, self.port))
            sock.sendall((command + '\n').encode('utf-8'))
            buffer = ""
            # 增量解码器：sock.recv 可能在多字节(中文)字符中间切断，逐片 data.decode
            # 会崩("unexpected end of data")。用它把不完整的尾字节留到下一片再拼。
            _decoder = codecs.getincrementaldecoder('utf-8')()
            file_path = None
            while True:
                try:
                    data = sock.recv(8192)
                    if not data:
                        break
                    decoded = _decoder.decode(data)
                    buffer += decoded
                    while '\n' in buffer:
                        line, buffer = buffer.split('\n', 1)
                        line = line.rstrip('\r')
                        if line.startswith('!FILE:'):
                            parts = line.split(':', 2)
                            if len(parts) == 3:
                                _, filename, b64_content = parts
                                if b64_content == 'FILE_NOT_FOUND':
                                    print(f"[!] 文件 {filename} 在服务端不存在")
                                elif b64_content.startswith('ERROR:'):
                                    print(f"[!] 文件 {filename} 传输错误: {b64_content}")
                                else:
                                    try:
                                        file_data = base64.b64decode(b64_content)

                                        # 检查是否为压缩文件（.gz 后缀）
                                        if filename.endswith('.gz'):
                                            print(f"[DECOMPRESS] 解压缩文件: {filename}")
                                            file_data = gzip.decompress(file_data)
                                            filename = filename[:-3]  # 去掉 .gz 后缀
                                            print(f"[DECOMPRESS] 解压后大小: {len(file_data) / 1024 / 1024:.2f} MB")

                                        with open(filename, 'wb') as f:
                                            f.write(file_data)
                                        abs_path = os.path.abspath(filename)
                                        self._remote_log(f"[FILE] 已接收并保存: {abs_path}")
                                        file_path = abs_path
                                    except Exception as e:
                                        print(f"[!] 解码文件 {filename} 失败: {e}")
                        elif line.startswith('!SLOG:'):
                            # 服务器回传的本 case 完整操作日志，写入 souren_execution.log
                            # （这些行原本只在服务器控制台，客户端日志之前拿不到）
                            self._remote_log(line[6:], echo=False)
                        else:
                            if line.strip():
                                self._remote_log(f"[INFO] {line}")
                except socket.timeout:
                    break
                except (ConnectionResetError, BrokenPipeError):
                    break
            return file_path
        except Exception as e:
            print(f"[!] 接收文件失败: {e}")
            return None
        finally:
            if sock:
                sock.close()

    def _receive_all_files_from_command(self, command, timeout=120):
        sock = None
        files = []
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(timeout)
            sock.connect((self.host, self.port))
            sock.sendall((command + '\n').encode('utf-8'))
            buffer = ""
            # 增量解码器：sock.recv 可能在多字节(中文)字符中间切断，逐片 data.decode
            # 会崩("unexpected end of data")。用它把不完整的尾字节留到下一片再拼。
            _decoder = codecs.getincrementaldecoder('utf-8')()
            while True:
                try:
                    data = sock.recv(8192)
                    if not data:
                        break
                    decoded = _decoder.decode(data)
                    buffer += decoded
                    while '\n' in buffer:
                        line, buffer = buffer.split('\n', 1)
                        line = line.rstrip('\r')
                        if line.startswith('!FILE:'):
                            parts = line.split(':', 2)
                            if len(parts) == 3:
                                _, filename, b64_content = parts
                                if b64_content == 'FILE_NOT_FOUND':
                                    print(f"[!] 文件 {filename} 在服务端不存在")
                                elif b64_content.startswith('ERROR:'):
                                    print(f"[!] 文件 {filename} 传输错误: {b64_content}")
                                else:
                                    try:
                                        file_data = base64.b64decode(b64_content)

                                        # 检查是否为压缩文件（.gz 后缀）
                                        if filename.endswith('.gz'):
                                            print(f"[DECOMPRESS] 解压缩文件: {filename}")
                                            file_data = gzip.decompress(file_data)
                                            filename = filename[:-3]  # 去掉 .gz 后缀
                                            print(f"[DECOMPRESS] 解压后大小: {len(file_data) / 1024 / 1024:.2f} MB")

                                        with open(filename, 'wb') as f:
                                            f.write(file_data)
                                        abs_path = os.path.abspath(filename)
                                        self._remote_log(f"[FILE] 已接收并保存: {abs_path}")
                                        files.append(abs_path)
                                    except Exception as e:
                                        print(f"[!] 解码文件 {filename} 失败: {e}")
                        elif line.startswith('!SLOG:'):
                            # 服务器回传的本 case 完整操作日志，写入 souren_execution.log
                            # （这些行原本只在服务器控制台，客户端日志之前拿不到）
                            self._remote_log(line[6:], echo=False)
                        else:
                            if line.strip():
                                self._remote_log(f"[INFO] {line}")
                except socket.timeout:
                    break
                except (ConnectionResetError, BrokenPipeError):
                    break
            return files
        except Exception as e:
            print(f"[!] 接收文件失败: {e}")
            return files
        finally:
            if sock:
                sock.close()

    def start_log(self, save_dir=None):
        if save_dir:
            os.makedirs(save_dir, exist_ok=True)
            original_cwd = os.getcwd()
            os.chdir(save_dir)
        else:
            original_cwd = None
        try:
            success = self._send_command_and_check_ok("DIAG:start_log", "[OK] 已启动")
            if success:
                print("[OK] 远程日志启动成功")
            else:
                print("[WARN] 远程日志启动失败")
            return success
        finally:
            if original_cwd:
                os.chdir(original_cwd)

    def stop_log(self):
        print("[*] 请求停止抓取并获取所有日志...")
        files = self._receive_all_files_from_command("DIAG:stop_log")
        signal_file = files[0] if len(files) > 0 else None
        scpi_file = files[1] if len(files) > 1 else None
        return signal_file, scpi_file

    def pvt_screenshot(self, save_dir=None):
        if save_dir:
            os.makedirs(save_dir, exist_ok=True)
            original_cwd = os.getcwd()
            os.chdir(save_dir)
        else:
            original_cwd = None
        try:
            file_path = self._receive_file_from_command("DIAG:pvt", timeout=60)
            return file_path
        finally:
            if original_cwd:
                os.chdir(original_cwd)

    def configure_log_and_restart(self, params, save_dir=None, timeout=60):
        """完整走一遍网页操作:点 Log -> Config -> 按 params 设置 Level/checkbox -> Apply
        -> 确认"Apply Success, Do you want to restart?"弹窗(点 OK)。
        params: dict，对应 souren_config.LOG_LEVEL_PARAMS(log_level + 各 checkbox 项)。
        成功返回 True；失败时若服务端回传了调试截图/页面源码，会保存到 save_dir 并返回 False。"""
        original_cwd = None
        if save_dir:
            os.makedirs(save_dir, exist_ok=True)
            original_cwd = os.getcwd()
            os.chdir(save_dir)
        sock = None
        try:
            params_b64 = base64.b64encode(json.dumps(params).encode('utf-8')).decode('ascii')
            command = f"DIAG:configure_log_and_restart?params={params_b64}"
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(timeout)
            sock.connect((self.host, self.port))
            sock.sendall((command + '\n').encode('utf-8'))
            buffer = ""
            _decoder = codecs.getincrementaldecoder('utf-8')()
            success = False
            while True:
                try:
                    data = sock.recv(8192)
                    if not data:
                        break
                    decoded = _decoder.decode(data)
                    buffer += decoded
                    while '\n' in buffer:
                        line, buffer = buffer.split('\n', 1)
                        line = line.rstrip('\r')
                        if line.startswith('!FILE:'):
                            parts = line.split(':', 2)
                            if len(parts) == 3:
                                _, filename, b64_content = parts
                                if b64_content in ('FILE_NOT_FOUND',) or b64_content.startswith('ERROR:'):
                                    print(f"[!] 调试文件 {filename} 获取失败: {b64_content}")
                                else:
                                    try:
                                        file_data = base64.b64decode(b64_content)
                                        if filename.endswith('.gz'):
                                            file_data = gzip.decompress(file_data)
                                            filename = filename[:-3]
                                        with open(filename, 'wb') as f:
                                            f.write(file_data)
                                        self._remote_log(f"[FILE] 调试文件已保存: {os.path.abspath(filename)}")
                                    except Exception as e:
                                        print(f"[!] 解码调试文件 {filename} 失败: {e}")
                        elif line.strip():
                            self._remote_log(line)
                            if '[OK]' in line:
                                success = True
                except socket.timeout:
                    break
                except (ConnectionResetError, BrokenPipeError):
                    break
            return success
        except Exception as e:
            print(f"[!] 配置日志级别并重启失败: {e}")
            return False
        finally:
            if sock:
                sock.close()
            if original_cwd:
                os.chdir(original_cwd)

    def _execute_command_and_receive_output(self, command, ready_string, timeout=180):
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(1.0)
        try:
            sock.connect((self.host, self.port))
            sock.sendall((command + '\n').encode('utf-8'))
            buffer = ""
            # 增量解码器：sock.recv 可能在多字节(中文)字符中间切断，逐片 data.decode
            # 会崩("unexpected end of data")。用它把不完整的尾字节留到下一片再拼。
            _decoder = codecs.getincrementaldecoder('utf-8')()
            output_lines = []
            ready = False
            start_time = time.time()
            while time.time() - start_time < timeout:
                try:
                    data = sock.recv(8192)
                    if not data:
                        break
                    decoded = _decoder.decode(data)
                    buffer += decoded
                    while '\n' in buffer:
                        line, buffer = buffer.split('\n', 1)
                        line = line.rstrip('\r')
                        output_lines.append(line)
                        if line.strip():
                            self._remote_log(line, echo=False)
                        if ready_string and ready_string in line:
                            ready = True
                except socket.timeout:
                    continue
                except (ConnectionResetError, BrokenPipeError):
                    break
                if ready:
                    time.sleep(2)
                    break
            sock.settimeout(2)
            while True:
                try:
                    data = sock.recv(8192)
                    if not data:
                        break
                    decoded = _decoder.decode(data)
                    buffer += decoded
                    while '\n' in buffer:
                        line, buffer = buffer.split('\n', 1)
                        line = line.rstrip('\r')
                        output_lines.append(line)
                        if line.strip():
                            self._remote_log(line, echo=False)
                except socket.timeout:
                    break
            return ready, output_lines
        except Exception as e:
            print(f"[!] 执行命令失败: {e}")
            return False, []
        finally:
            sock.close()

    @staticmethod
    def ping(host, port):
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(3)
            sock.connect((host, port))
            sock.close()
            return True
        except:
            return False

    def collect_current_logs(self, save_dir=None):
        """打包回传 /tmp/yc1100/current 下基站软件自身维护的全部日志文件
        (mainctrl.log / NR_<ts>.log 等，网页 Restart 后由基站软件自己写入)。
        直接解包到 save_dir(与结果文件同目录)，不再新建 gnb_current_logs 子文件夹。"""
        if save_dir is None:
            save_dir = os.getcwd()
        os.makedirs(save_dir, exist_ok=True)
        original_cwd = os.getcwd()
        os.chdir(save_dir)
        try:
            print(f"[*] 开始收集基站 current 日志，保存目录: {save_dir}")
            archive = self._receive_file_from_command("DIAG:collect_current_logs", timeout=120)
            if not archive:
                print("[WARN] 未收到基站 current 日志压缩包")
                return None
            try:
                with tarfile.open(archive) as tf:
                    names = [m.name.lstrip('./') for m in tf.getmembers() if m.isfile()]
                    tf.extractall(save_dir)
                print(f"[OK] 基站 current 日志已解包 {len(names)} 个文件: {', '.join(sorted(names))}")
            finally:
                try:
                    os.unlink(archive)
                except Exception:
                    pass
            return save_dir
        finally:
            os.chdir(original_cwd)

    def collect_core_logs(self, save_dir=None):
        # 收集核心网(mt5gc)全部日志：服务端 tar 打包一次性回传，客户端解包到
        # save_dir/core_network_logs。整包传输 + gzip 压缩，比逐文件传快得多。
        if save_dir is None:
            save_dir = os.getcwd()
        target_dir = os.path.join(save_dir, "core_network_logs")
        os.makedirs(target_dir, exist_ok=True)
        original_cwd = os.getcwd()
        os.chdir(target_dir)
        try:
            print(f"[*] 开始收集核心网日志，保存目录: {target_dir}")
            archive = self._receive_file_from_command("DIAG:collect_core_logs", timeout=180)
            if not archive:
                print("[WARN] 未收到核心网日志压缩包")
                return None
            try:
                with tarfile.open(archive) as tf:
                    names = [m.name.lstrip('./') for m in tf.getmembers() if m.isfile()]
                    tf.extractall(target_dir)
                print(f"[OK] 核心网日志已解包 {len(names)} 个文件: {', '.join(sorted(names))}")
            finally:
                try:
                    os.unlink(archive)
                except Exception:
                    pass
            return target_dir
        finally:
            os.chdir(original_cwd)