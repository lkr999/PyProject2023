from gekko import GEKKO
m = GEKKO()            # create GEKKO model
x = m.Var()            # define new variable, default=0
y = m.Var()            # define new variable, default=0
m.Equations([3*x+2*y==1, x+2*y==0])  # equations
m.solve(disp=False)    # solve
print(x.value,y.value) # print solution

# usage: thermo('mw') for constants
# # thermo('lvp',T) for temperature dependent
# from gekko import GEKKO, chemical
# m = GEKKO()
# c = chemical.Properties(m)
# # add compounds
# c.compound('water')
# c.compound('hexane')
# c.compound('heptane')
# # molecular weight
# mw = c.thermo('mw')
# # liquid vapor pressure
# T = m.Param(value=310)
# vp = c.thermo('lvp',T)
# m.solve(disp=False)
# print(mw)
# print(vp)