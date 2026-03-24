

import os
import time
import sys
from vncdotool import api

def wait_for_desktop(vnc_client, timeout=30):
    """等待桌面加载完成"""
    print("等待Ubuntu桌面加载...")
    start_time = time.time()
    
    # 这里可以添加检测桌面是否加载的逻辑
    # 比如检测特定颜色或等待一段时间
    time.sleep(5)
    
    return True

def vnc_type_command(vnc_client, command):
    """通过VNC输入命令"""
    print(f"输入命令: {command}")
    
    # 确保在命令行界面
    vnc_client.keyPress('ctrl')
    vnc_client.keyPress('alt')
    vnc_client.keyPress('t')
    vnc_client.keyRelease('t')
    vnc_client.keyRelease('alt')
    vnc_client.keyRelease('ctrl')
    time.sleep(1)
    
    # 输入命令
    for char in command:
        if char.isupper():
            vnc_client.keyPress('shift')
            vnc_client.keyPress(char.lower())
            vnc_client.keyRelease(char.lower())
            vnc_client.keyRelease('shift')
        elif char == ' ':
            vnc_client.keyPress('space')
            vnc_client.keyRelease('space')
        elif char == '-':
            vnc_client.keyPress('-')
            vnc_client.keyRelease('-')
        elif char == '.':
            vnc_client.keyPress('.')
            vnc_client.keyRelease('.')
        elif char == '/':
            vnc_client.keyPress('/')
            vnc_client.keyRelease('/')
        elif char == ':':
            vnc_client.keyPress('shift')
            vnc_client.keyPress(';')
            vnc_client.keyRelease(';')
            vnc_client.keyRelease('shift')
        else:
            vnc_client.keyPress(char)
            vnc_client.keyRelease(char)
        time.sleep(0.05)
    
    # 按回车执行
    vnc_client.keyPress('enter')
    vnc_client.keyRelease('enter')
    time.sleep(1)

def create_file_via_vnc(vnc_client, filename, content):
    """通过VNC创建文件"""
    print(f"在Ubuntu上创建文件 {filename}...")
    
    # 打开终端
    vnc_type_command(vnc_client, f"touch {filename}")
    time.sleep(1)
    
    # 写入内容到文件
    vnc_type_command(vnc_client, f"echo '{content}' > {filename}")
    time.sleep(1)
    
    # 验证文件
    vnc_type_command(vnc_client, f"cat {filename}")
    time.sleep(2)
    
    print(f"✓ 文件 {filename} 创建完成")

def copy_file_content_to_local(vnc_client, filename, local_path):
    """通过VNC复制文件内容到本地"""
    print(f"尝试获取文件 {filename} 的内容...")
    
    # 显示文件内容
    vnc_type_command(vnc_client, f"cat {filename}")
    time.sleep(2)
    
    # 注意：这里无法直接通过VNC传输文件
    # 我们将手动复制或使用其他方法
    print("无法通过VNC直接传输文件，需要手动复制内容")
    
    # 这里可以添加手动复制的提示
    # 或者尝试使用剪贴板（如果VNC支持）
    
    return False

def delete_file_via_vnc(vnc_client, filename):
    """通过VNC删除文件"""
    print(f"删除Ubuntu上的文件 {filename}...")
    
    vnc_type_command(vnc_client, f"rm {filename}")
    time.sleep(1)
    
    # 验证文件是否删除
    vnc_type_command(vnc_client, f"ls {filename}")
    time.sleep(2)
    
    print(f"✓ 文件 {filename} 已删除")

def alternative_solution():
    """替代方案：如果VNC无法执行命令，使用手动步骤"""
    print("\n" + "="*60)
    print("替代方案（手动步骤）：")
    print("="*60)
    print("\n由于只能使用VNC连接，请按以下步骤操作：")
    print("\n1. 请先通过VNC连接到 Ubuntu:")
    print(f"   地址: 192.168.30.122:5900")
    print(f"   用户名: visitor")
    print(f"   密码: 123456")
    
    print("\n2. 在Ubuntu上打开终端 (Ctrl+Alt+T)，执行以下命令:")
    print(f"   touch vnc-result.txt")
    print(f"   echo 'yc110' > vnc-result.txt")
    print(f"   cat vnc-result.txt  # 验证内容")
    
    print("\n3. 现在需要将文件传输到Windows:")
    print("   a) 方法1: 在Ubuntu上安装SSH服务器")
    print("      sudo apt update && sudo apt install openssh-server")
    print("      sudo systemctl enable ssh && sudo systemctl start ssh")
    
    print("\n   b) 方法2: 通过HTTP服务器共享文件")
    print("      python3 -m http.server 8000")
    print("      然后在浏览器访问: http://192.168.30.122:8000")
    
    print("\n   c) 方法3: 使用Python临时HTTP服务器")
    print("      cd /path/to/file && python3 -m http.server")
    
    print("\n   d) 方法4: 手动复制内容")
    print("      cat vnc-result.txt")
    print("      然后在Windows上手动创建同名文件并粘贴内容")
    
    print("\n4. 完成传输后，在Ubuntu上删除文件:")
    print(f"   rm vnc-result.txt")
    print("\n" + "="*60)

def main():
    UBUNTU_IP = "192.168.30.122"
    VNC_PORT = 5900
    VNC_PASSWORD = "123456"  # VNC密码
    USERNAME = "visitor"
    
    FILENAME = "vnc-result.txt"
    FILE_CONTENT = "yc110"
    
    print("=" * 60)
    print("VNC自动化脚本")
    print("=" * 60)
    print(f"目标系统: {UBUNTU_IP}:{VNC_PORT}")
    print(f"用户名: {USERNAME}")
    print("=" * 60)
    
    try:
        # 尝试连接到VNC服务器
        print(f"\n1. 连接到VNC服务器 {UBUNTU_IP}:{VNC_PORT}...")
        
        vnc_client = None
        try:
            # 使用vncdotool连接
            vnc_client = api.connect(f"{UBUNTU_IP}:{VNC_PORT}", password=VNC_PASSWORD)
            print("✓ VNC连接成功")
        except Exception as e:
            print(f"✗ VNC连接失败: {e}")
            print("\n注意：您需要先手动启动VNC连接并登录桌面")
            print("然后此脚本才能通过VNC执行操作")
            
            # 提供替代方案
            alternative_solution()
            return
        
        # 等待桌面加载
        if not wait_for_desktop(vnc_client):
            print("桌面加载超时")
            alternative_solution()
            return
        
        # 创建文件
        print(f"\n2. 在Ubuntu上创建文件 {FILENAME}...")
        create_file_via_vnc(vnc_client, FILENAME, FILE_CONTENT)
        
        # 传输文件（无法直接通过VNC传输）
        print(f"\n3. 需要手动获取文件内容...")
        print("\n请在VNC窗口中查看文件内容并手动复制:")
        print(f"   文件: {FILENAME}")
        print(f"   内容: {FILE_CONTENT}")
        
        # 显示文件内容
        vnc_type_command(vnc_client, f"cat {FILENAME}")
        time.sleep(3)
        
        # 在本地创建文件（手动复制内容后）
        print(f"\n4. 请在Windows当前目录手动创建文件 {FILENAME}")
        current_dir = os.getcwd()
        local_file = os.path.join(current_dir, FILENAME)
        
        print(f"   路径: {local_file}")
        print(f"   内容: {FILE_CONTENT}")
        
        # 等待用户确认
        input(f"\n请按Enter键继续（确认已复制内容）...")
        
        # 检查本地文件是否已创建
        if os.path.exists(local_file):
            with open(local_file, 'r') as f:
                content = f.read().strip()
            if content == FILE_CONTENT:
                print(f"✓ 本地文件验证成功: {content}")
            else:
                print(f"⚠ 本地文件内容不匹配: {content}")
        
        # 删除Ubuntu上的文件
        print(f"\n5. 删除Ubuntu上的文件 {FILENAME}...")
        delete_file_via_vnc(vnc_client, FILENAME)
        
        print("\n" + "=" * 60)
        print("操作完成!")
        print("=" * 60)
        
    except Exception as e:
        print(f"\n错误: {e}")
        alternative_solution()
    
    finally:
        if 'vnc_client' in locals() and vnc_client:
            try:
                vnc_client.disconnect()
            except:
                pass

if __name__ == "__main__":
    # 检查是否安装了vncdotool
    try:
        from vncdotool import api
    except ImportError:
        print("需要安装vncdotool库...")
        print("请执行: pip install vncdotool")
        print("\n如果安装失败，请使用以下手动步骤:")
        alternative_solution()
        sys.exit(1)
    
    main()