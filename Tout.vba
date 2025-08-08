Function Tout(Tin, Tsoil, mFlow, L, z, P, eps, D1, D2, D3, D4)
' Calculates the temperature drop through a pipe segment
'   to find outlet temperature at the end of the pipe
'   by Tol,Hakan Ibrahim from the research  carried out in SALK GeoWatt project at VITO NV and at TU/e

' INPUTS
'    Tin      : The temperature at the pipe inlet   [°C]
'    Tsoil    : Undisturbed soil temperature        [°C]
'    mFlow    : Mass flow rate                      [kg/s]
'    L        : Pipe length                         [m]
'    z        : Pipe depth                          [m]
'    P        : Static Pressure (10,6 - supply, return) [bar]
'    D1-4     : Diameters for pipe,insulation,outer casing, and outside [m] or [mm]

' THERMAL PROPERTIES
' Water
rho = rhoL_T(Tin)         '[kg/m3]
Cp = CpL_T(Tin) * 1000    '[J/kgK]
TCw = -0.00000967398 * Tin ^ 2 + 0.00215246 * Tin + 0.55748 '[W/mK]
vis = my_pT(P, Tin)

' Pipe Layers
TCp = 51    '[W/mK]
TCi = 0.027 '[W/mK]
TCc = 0.43  '[W/mK]
TCs = 1.6   '[W/mK] - Soil

' THERMAL RESISTANCES
' Pipe Layers
Rp = Log(D2 / D1) / (2 * PiNumber() * L * TCp)
Ri = Log(D3 / D2) / (2 * PiNumber() * L * TCi)
Rc = Log(D4 / D3) / (2 * PiNumber() * L * TCc)

' Soil
Rs = Log((z / (D4 / 2)) + ((z / (D4 / 2)) ^ 2 - 1) ^ 0.5) / (2 * PiNumber() * L * TCs)

' Inner Flow

f = f_Clamond(Reynolds(mFlow, D1, Tin, P), eps / D1)
Pr = Cp * vis / TCw

Nu = ((f / 8) * (Re - 1000) * Pr) / (1 + 12.7 * (f / 8) ^ 0.5 * (Pr ^ (2 / 3) - 1))
hf = Nu * TCw / (D1 / 1000)
Af = PiNumber() * D1 * L

Rf = 1 / (hf * Af)

Rt = Rf + Rp + Ri + Rc + Rs

Tout = Tsoil + (Tin - Tsoil) * Exp(-(1 / (Rt * mFlow * Cp)))

End Function
