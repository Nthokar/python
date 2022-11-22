from matplotlib.ticker import FuncFormatter
import matplotlib.pyplot as plt
import numpy as np

x = np.arange(4)
money = [1.5e5, 2.5e6, 5.5e6, 2.0e7]


def millions(x, pos):
    'The two args are the value and tick position'
    return '$%1.1fM' % (x*1e-6)

formatter = FuncFormatter(millions)

fig, ax = plt.subplots()
ax.yaxis.grid(True, linestyle='--', which='major',
               color='grey', alpha=.25)
ax.yaxis.set_major_formatter(formatter)
plt.bar(x, money, width=0.2)
plt.bar(x, money, width=0.2)
plt.xticks(x, ('Bill', 'Fred', 'Mary', 'Sue'))
plt.show()