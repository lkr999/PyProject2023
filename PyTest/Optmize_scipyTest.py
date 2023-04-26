import numpy as np
import matplotlib.pyplot as plt
import itertools

import xlwings as xl
from pymsgbox import alert, confirm, prompt

from scipy.optimize import minimize, LinearConstraint, Bounds



def fun_object(x):
    return x[0]**2 + x[0]*x[1]

def eq_constraint(x):
    return x[0]**3 + x[0]*x[1] - 100

def ineq_constraint2(x):
    return 50 - np.sum(x) 

def ineq_constraint(x):
    return x[0]**2 + x[1] - 50

# np.sum([Ri / 100 * g_Pdty[i] * (g_Price[i] - g_VC[i] - g_SC[i]) - FV for i, Ri in enumerate(Sales_R)])

def fun_cal():
    bounds = [[-100,100], [-100,100]]
    constraint_1 = {'type': 'eq', 'fun': eq_constraint}
    constraint_2 = {'type': 'ineq', 'fun': ineq_constraint}
    constraint_3 = {'type': 'ineq', 'fun': ineq_constraint2}
    # constraint_3 = {'type': 'ineq', 'fun': ieq_constraint}
    constraint = [constraint_1, constraint_2, constraint_3]
    x0 = [1,1]

    result = minimize(fun_object, x0, method='SLSQP', bounds=bounds, constraints=constraint)

    print(result.fun, result.x)
    print(result)

fun_cal() 