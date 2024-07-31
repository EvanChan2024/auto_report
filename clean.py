import numpy as np
from numba import njit


@njit
def sliding_window_median_absolute_deviation(data, window):
    n = len(data)
    half_window = window // 2
    median_list = np.full(n, np.nan)
    mad_list = np.full(n, np.nan)
    c = 1.4826  # 常数因子

    # 为了加速计算，使用预分配的数组
    buffer = np.empty(window, dtype=float)
    buffer.fill(np.nan)

    for i in range(n):
        # 窗口范围
        start = max(0, i - half_window)
        end = min(n, i + half_window)  # 这里有问题
        k = end - start

        # 填充缓冲区
        buffer[:k] = data[start:end]

        # 处理缓冲区未满的部分
        if k < window:
            buffer[k:] = np.nan

        # 计算窗口内的中位数和MAD
        if k > 0:  # 确保窗口内有数据
            median = np.nanmedian(buffer[:k])
            mad = np.nanmedian(np.abs(buffer[:k] - median))
            median_list[i] = median
            mad_list[i] = c * mad

        # print(f"{i}: Window [{start}:{end}] with k={k}")  # 调试代码：显示窗口范围和k值
    return median_list, mad_list


def GPSDataFilter2(data, threshold, rateThreshold, window, thresholdFactor):
    data = np.array(data)

    # 如果所有数据都是 NaN，直接返回原数据
    if np.all(np.isnan(data)):
        return data

    # 将超过 threshold 的数据设置为 NaN
    data = np.where(np.abs(data) > threshold, np.nan, data)
    # 计算数据的差分
    ddata = np.diff(data)
    # 找到绝对差分大于 rateThreshold 的索引
    outInddy0 = np.where(np.abs(ddata) > rateThreshold)[0]
    # 将这些索引及其下一个索引合并，并去重
    outInd = np.unique(np.concatenate([outInddy0, outInddy0 + 1]))
    # 将这些索引处的数据设置为 NaN
    data[outInd] = np.nan

    ###
    # 使用滑动窗口方法计算局部中位数和换算 MAD
    median, mad = sliding_window_median_absolute_deviation(data, window)

    # 填充窗口边缘
    pad_size = (len(data) - len(median)) // 2
    median = np.pad(median, (pad_size, pad_size), 'edge')
    mad = np.pad(mad, (pad_size, pad_size), 'edge')

    # 确保 median 和 mad 与 data 的大小匹配
    if len(median) != len(data):
        median = np.pad(median, (0, len(data) - len(median)), 'edge')
    if len(mad) != len(data):
        mad = np.pad(mad, (0, len(data) - len(mad)), 'edge')

    # 检测异常值
    youtInd = np.abs(data - median) > (thresholdFactor * mad)
    # 找到非异常值的索引
    outNum = np.where(youtInd)[0]

    # 如果没有异常值，直接返回数据
    if len(outNum) == 0:
        return data

    # 创建一个用于存储异常值区间的列表
    rngV = [outNum[0]]
    for i in range(len(outNum) - 1):
        if outNum[i + 1] - outNum[i] != 1:
            rngV.extend([outNum[i], outNum[i + 1]])
    rngV.append(outNum[-1])

    # 将异常值区间重新整理为二维数组
    rng = np.array(rngV).reshape(-1, 2)

    # 遍历每个异常值区间
    for start, end in rng:
        # 如果区间前后均为 NaN，则将该区间设置为 NaN
        if np.isnan(data[max(start - 1, 0)]) and np.isnan(data[min(end + 1, len(data) - 1)]):
            data[start:end + 1] = np.nan
    ###

    # 计算数据的绝对值
    dataAbs = np.abs(data)
    # 计算绝对值的 99.5 百分位数
    d99 = np.percentile(dataAbs, 99.5)
    # 找到绝对值大于 99.5 百分位数的索引
    ind = np.where(dataAbs > d99)[0]
    # 计算这些索引的差分
    dind = np.diff(ind)
    # 找到差分大于 1 的索引
    dpatind = np.where(dind > 1)[0]

    # 如果没有找到这样的索引，直接返回数据
    if len(dpatind) == 0:
        return data

    # 创建一个用于存储间隔区间的二维列表
    dM = [[ind[0], ind[dpatind[0]]]]
    for i in range(len(dpatind) - 1):
        dM.append([ind[dpatind[i] + 1], ind[dpatind[i + 1]]])
    dM.append([ind[dpatind[-1] + 1], ind[-1]])

    # 遍历每个间隔区间
    for ind1, ind2 in dM:
        # 如果区间前后均为 NaN，则将该区间设置为 NaN
        if np.isnan(data[max(ind1 - 1, 0)]) and np.isnan(data[min(ind2 + 1, len(data) - 1)]):
            data[ind1:ind2 + 1] = np.nan

    # 返回处理后的数据
    return data

