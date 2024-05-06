import numpy as np
import matplotlib.pyplot as plt

# 假设loss数据按行存放在一个文本文件中
with open('yolov8x.txt', 'r') as file:
    lines = file.readlines()
    # 假设每行是一个浮点数，直接转换为float类型
    loss_history = [float(line.strip()) for line in lines]

# 转换为numpy数组方便后续处理
loss_history = np.array(loss_history)

# 绘制loss曲线
num_epochs = len(loss_history)
plt.figure(figsize=(10, 6))
plt.plot(range(1, num_epochs + 1), loss_history, marker='o', markersize=1)
plt.xlabel('Iteration')
plt.ylabel('Loss')
plt.title('Training Loss Curve')
plt.grid(True)
plt.show()
