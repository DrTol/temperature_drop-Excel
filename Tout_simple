Function Tout_simple(Tin, mFlow, L, DN)
'' Simplified version of the above function

' Catalogue Data (LOGSTOR - Steel & Insulation Series 1)
d_DN = VBA.Array(20, 25, 32, 40, 50, 65, 80, 100, 125, 150, 200, 250, 300, 350, 400, 450, 500, 600, 700, 800, 900, 1000, 1100, 1200)
d_D1 = VBA.Array(21.7, 28.5, 37.2, 43.1, 54.5, 70.3, 82.5, 107.1, 132.5, 160.3, 210.1, 263, 312.7, 344.4, 393.8, 444.4, 495.4, 595.8, 695, 795.4, 894, 994, 1096, 1194)
d_D2 = VBA.Array(26.9, 33.7, 42.4, 48.3, 60.3, 76.1, 88.9, 114.3, 139.7, 168.3, 219.1, 273, 323.9, 355.6, 406.4, 457, 508, 610, 711, 813, 914, 1016, 1118, 1219)
d_D3 = VBA.Array(84, 84, 104, 104, 119, 134, 154, 193.6, 218.2, 242.8, 306.8, 390.4, 439.6, 488.8, 548.6, 618, 696.8, 784.4, 882.6, 981.2, 1079.6, 1178, 1276.4, 1375)
d_D4 = VBA.Array(90, 90, 110, 110, 125, 140, 160, 200, 225, 250, 315, 40, 450, 500, 560, 630, 710, 800, 900, 1000, 1100, 1200, 1300, 1400)

pos = Application.Match(DN, d_DN, False)

D1 = d_D1(pos - 1)
D2 = d_D2(pos - 1)
D3 = d_D3(pos - 1)
D4 = d_D4(pos - 1)
    
z = 1000    ' [mm] Pipe depth
Tsoil = 10  ' [°C] Undisturbed soil temperature
P = 8       ' [bar] Water pressure

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

Tout_simple = Tsoil + (Tin - Tsoil) * Exp(-(1 / (Rt * mFlow * Cp)))

End Function
