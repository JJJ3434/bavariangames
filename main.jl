module BG
using XLSX
using DataFrames

Gruppen = DataFrame(XLSX.readtable("BavarianGames.xlsx", "Gruppe"))

PktMasskrugschieben = DataFrame(XLSX.readtable("BavarianGames.xlsx", "PktMasskrugschieben"))
PktWackelturm = DataFrame(XLSX.readtable("BavarianGames.xlsx", "PktWackelturm"))
PktHolzscheitelwurf = DataFrame(XLSX.readtable("BavarianGames.xlsx", "PktHolzscheitelwurf"))
PktReifenwuchten = DataFrame(XLSX.readtable("BavarianGames.xlsx", "PktReifenwuchten"))
PktBulldogziehen = DataFrame(XLSX.readtable("BavarianGames.xlsx", "PktBulldogziehen"))

end
