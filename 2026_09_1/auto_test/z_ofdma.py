import numpy as np
import matplotlib.pyplot as plt

# ==================== 仿真参数 ====================
num_symbols = 1000      # OFDM符号数
fft_size = 64           # FFT点数
cp_len = 16             # 循环前缀长度
mcs_type = '16QAM'      # 调制类型

# ==================== 辅助函数 ====================
def generate_qam_symbols(num, mod_type):
    """生成QAM调制符号"""
    if mod_type == 'QPSK':
        const = (np.array([1+1j, 1-1j, -1+1j, -1-1j]) / np.sqrt(2))
    elif mod_type == '16QAM':
        const = (np.array([-3-3j, -3-1j, -3+3j, -3+1j,
                           -1-3j, -1-1j, -1+3j, -1+1j,
                           3-3j, 3-1j, 3+3j, 3+1j,
                           1-3j, 1-1j, 1+3j, 1+1j]) / np.sqrt(10))
    else:
        raise ValueError("Unsupported modulation")
    idx = np.random.randint(0, len(const), num)
    return const[idx]

def calculate_papr(signal):
    """计算PAPR (dB)"""
    power = np.abs(signal)**2
    peak_power = np.max(power)
    avg_power = np.mean(power)
    papr = 10 * np.log10(peak_power / avg_power)
    return papr

# ==================== CP-OFDM 发射机 ====================
def cp_ofdm_tx(data_symbols, fft_size, cp_len):
    """CP-OFDM发射机"""
    # 子载波映射（这里假设所有子载波均用于数据）
    freq_domain = data_symbols
    # IFFT
    time_domain = np.fft.ifft(freq_domain, fft_size)
    # 添加CP
    cp = time_domain[-cp_len:]
    tx_signal = np.concatenate([cp, time_domain])
    return tx_signal

# ==================== DFT-s-OFDM 发射机 ====================
def dft_s_ofdm_tx(data_symbols, fft_size, cp_len):
    """DFT-s-OFDM发射机"""
    # DFT扩频
    freq_domain = np.fft.fft(data_symbols)
    # 子载波映射
    mapped_symbols = np.zeros(fft_size, dtype=complex)
    mapped_symbols[:len(freq_domain)] = freq_domain
    # IFFT
    time_domain = np.fft.ifft(mapped_symbols, fft_size)
    # 添加CP
    cp = time_domain[-cp_len:]
    tx_signal = np.concatenate([cp, time_domain])
    return tx_signal

# ==================== 主仿真流程 ====================
def run_simulation():
    papr_cp = []
    papr_dft = []
    for _ in range(num_symbols):
        # 生成数据符号
        data = generate_qam_symbols(fft_size, mcs_type)
        # CP-OFDM发射
        tx_cp = cp_ofdm_tx(data, fft_size, cp_len)
        papr_cp.append(calculate_papr(tx_cp))
        # DFT-s-OFDM发射
        tx_dft = dft_s_ofdm_tx(data, fft_size, cp_len)
        papr_dft.append(calculate_papr(tx_dft))
    return np.array(papr_cp), np.array(papr_dft)

# 运行仿真
papr_cp, papr_dft = run_simulation()

# 绘制CCDF曲线
def plot_ccdf(papr_values, label):
    sorted_papr = np.sort(papr_values)
    ccdf = 1 - np.arange(1, len(sorted_papr)+1) / len(sorted_papr)
    plt.semilogy(sorted_papr, ccdf, label=label)

plt.figure(figsize=(10, 6))
plot_ccdf(papr_cp, f'CP-OFDM (Mean PAPR = {np.mean(papr_cp):.2f} dB)')
plot_ccdf(papr_dft, f'DFT-s-OFDM (Mean PAPR = {np.mean(papr_dft):.2f} dB)')
plt.xlabel('PAPR (dB)')
plt.ylabel('CCDF (Pr(PAPR > threshold))')
plt.title(f'PAPR Comparison: CP-OFDM vs DFT-s-OFDM ({mcs_type})')
plt.grid(True)
plt.legend()
plt.show()