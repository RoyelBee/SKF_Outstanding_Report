# library
import numpy as np
import matplotlib.pyplot as plt

# create data
Target = [1, 4, 6, 8]
Sales = [2, 4, 5, 5]
x = range(len(Target))
xx = range(len(Sales))


# Change the color and its transparency
plt.fill_between(x, Target, color="skyblue", alpha=1)
plt.plot(xx, Sales, color="red", linewidth=5, linestyle="-")
plt.xlabel('Days')
plt.ylabel('Sales Amount')
plt.legend(Sales)
plt.show()
