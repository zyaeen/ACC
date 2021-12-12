import matplotlib.pyplot as plt
import numpy as np

x = np.linspace(-2, 12, 100)
k = 24
f = lambda x: np.sin(k) * (x ** 2) + np.cos(2 * k) * x + np.log(1 / k) / 2 / k - 1 / k

y = f(x)

plt.plot(x, y)

roots = np.poly1d([np.sin(k), np.cos(2 * k), np.log(1 / k) / 2 / k - 1 / k]).roots
print(roots)

plt.plot(roots[0],
         0,
        marker='o',
        markersize=1,
             color='red')

plt.plot(roots[1],
         0,
        marker='o',
        markersize=1,
             color='red')

plt.show()
