# The parabolic PDE equation describes the evolution of temperature
#  for the interior region of the rod. This model is modified to make
#  one end of the device fixed and the other temperature at the end of the
#  device calculated.
import numpy as np
from gekko import GEKKO
import matplotlib.pyplot as plt
import matplotlib.animation as animation

# Steel temperature profile
# Diameter = 3 cm
# Length = 10 cm
seg = 100              # number of segments
T_melt = 1426             # melting temperature of H13 steel
pi = 3.14159          # pi
d = 3 / 100          # diameter (m)
L = 10 / 100         # length (m)
L_seg = L / seg          # length of a segment (m)
A = pi * d**2 / 4    # cross-sectional area (m)
As = pi * d * L_seg   # surface heat transfer area (m^2)
heff = 5.8              # heat transfer coeff (W/(m^2*K))
keff = 28.6             # thermal conductivity in H13 steel (W/m-K)
rho = 7760             # density of H13 steel (kg/m^3)
cp = 460              # heat capacity of H13 steel (J/kg-K)
Ts = 23               # temperature of the surroundings (°C)

m = GEKKO()  # create GEKKO model

tf = 3000
nt = int(tf/30) + 1
dist = np.linspace(0, L, seg+2)
m.time = np.linspace(0, tf, nt)
T1 = m.MV(ub=T_melt)        # temperature 1 (°C)
T1.value = np.ones(nt) * 23  # start at room temperature
T1.value[10:] = 80         # step at 300 sec

T = [m.Var(23) for i in range(seg)]  # temperature of the segments (°C)

T2 = m.MV(ub=T_melt)        # temperature 2 (°C)
T2.value = np.ones(nt) * 23  # start at room temperature
T2.value[50:] = 100         # step at 300 sec

# Energy balance for the segments
# accumulation =
#    (heat gained from upper segment)
#  - (heat lost to lower segment)
#  - (heat lost to surroundings)
# Units check
# kg/m^3 * m^2 * m * J/kg-K * K/sec =
#     W/m-K   * m^2 *  K / m
#  -  W/m-K   * m^2 *  K / m
#  -  W/m^2-K * m^2 *  K

# first segment
m.Equation(rho*A*L_seg*cp*T[0].dt() ==
           keff*A*(T1-T[0])/L_seg
           - keff*A*(T[0]-T[1])/L_seg
           - heff*As*(T[0]-Ts))
# middle segments
m.Equations([rho*A*L_seg*cp*T[i].dt() ==
             keff*A*(T[i-1]-T[i])/L_seg
             - keff*A*(T[i]-T[i+1])/L_seg
             - heff*As*(T[i]-Ts) for i in range(1, seg-1)])
# last segment
m.Equation(rho*A*L_seg*cp*T[-1].dt() ==
           keff*A*(T[-2]-T[-1])/L_seg
           - keff*A*(T[-1]-T2)/L_seg
           - heff*As*(T[-1]-Ts))

# simulation
m.options.IMODE = 4
m.solve()

# plot results
plt.figure()
tm = m.time / 60.0
plt.plot(tm, T1.value, 'k-', lw=2, label=r'$T_{left}\,(^oC)$')
plt.plot(tm, T[5].value, ':', color='yellow', label=r'$T_{5}\,(^oC)$')
plt.plot(tm, T[15].value, '--', color='red', label=r'$T_{15}\,(^oC)$')
plt.plot(tm, T[25].value, '--', color='green', label=r'$T_{50}\,(^oC)$')
plt.plot(tm, T[45].value, '-.', color='gray', label=r'$T_{85}\,(^oC)$')
plt.plot(tm, T[95].value, '-', color='orange', label=r'$T_{95}\,(^oC)$')
plt.plot(tm, T2.value, 'b-', lw=2, label=r'$T_{right}\,(^oC)$')
plt.ylabel(r'$T\,(^oC$)')
plt.xlabel('Time (min)')
plt.xlim([0, 50])
plt.legend(loc=4)
plt.savefig('heat.png')

# create animation as heat.mp4
fig = plt.figure(figsize=(5, 4))
fig.set_dpi(300)
ax1 = fig.add_subplot(1, 1, 1)

# store results in d
d = np.empty((seg+2, len(m.time)))
d[0] = np.array(T1.value)
for i in range(seg):
    d[i+1] = np.array(T[i].value)
d[-1] = np.array(T2.value)
d = d.T

k = 0


def animate(i):
    global k
    k = min(len(m.time)-1, k)
    ax1.clear()
    plt.plot(dist*100, d[k], color='red', label=r'Temperature ($^oC$)')
    plt.text(1, 100, 'Elapsed time '+str(round(m.time[k]/60, 2))+' min')
    plt.grid(True)
    plt.ylim([20, 110])
    plt.xlim([0, L*100])
    plt.ylabel(r'T ($^oC$)')
    plt.xlabel(r'Distance (cm)')
    plt.legend(loc=1)
    k += 1


anim = animation.FuncAnimation(fig, animate, frames=len(m.time), interval=20)
# requires ffmpeg to save mp4 file
#  available from https://ffmpeg.zeranoe.com/builds/
#  add ffmpeg.exe to path such as C:\ffmpeg\bin\ in
#  environment variables
try:
    anim.save('heat.mp4', fps=10)
except:
    print('requires ffmpeg to save mp4 file')
    plt.show()
