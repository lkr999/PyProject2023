
import numpy as np
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D

import xlwings as xw

from gekko import GEKKO

m = GEKKO()

# eq = m.Param(value=40)

x1, x2, x3, x4 = [m.Var() for i in range(4)]

x1.value = 1
x2.value = 5
x3.value = 5
x4.value = 1

x1.lower = 1
x2.lower = 1
x3.lower = 1
x4.lower = 1

x1.upper = 5
x2.upper = 5
x3.upper = 5
x4.upper = 5

m.Equation(x1*x2*x3*x4 >=25)
m.Equation(x1**2 + x2**2 + x3**2 + x4**2 == 40)

# Objective ---------------------
m.Obj(x1*x4*(x1+x2+x3) + x3)

# Set Global option -------------
m.options.IMODE = 3  # slowly state optimization

# Solve Simulation ----------------
m.solve()

# Result --------------------------
print(x1.value)
print(x2.value)
print(x3.value)
print(x4.value)


# from gekko import GEKKO
# import numpy as np
# m = GEKKO()
# f = m.Param(np.linspace(0,2,200))
# lower = m.Var()
#
# d=1e-5
# x_data = [-1e5,0.25,0.25+d,1,1+d,1.5,1.5+d,1e5]
# y_data = [0.1,0.1,0.75,0.75,1.25,1.25,1.7,1.7]
#
# m.pwl(f,lower,x_data,y_data)
# #m.Equation(f>=lower)
#
# m.options.IMODE = 2
# m.solve()
#
# import matplotlib.pyplot as plt
# plt.plot(f,lower,'b--')
# plt.plot(x_data,y_data,'ro')
# plt.xlim([0,2])
# plt.show()


