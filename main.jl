module BG
using XLSX
using DataFrames

Gruppen = DataFrame(XLSX.readtable("BavarianGames.xlsx", "Gruppe"))

PktMasskrugschieben = DataFrame(XLSX.readtable("BavarianGames.xlsx", "PktMasskrugschieben"))
PktWackelturm = DataFrame(XLSX.readtable("BavarianGames.xlsx", "PktWackelturm"))
PktHolzscheitelwurf = DataFrame(XLSX.readtable("BavarianGames.xlsx", "PktHolzscheitelwurf"))
PktReifenwuchten = DataFrame(XLSX.readtable("BavarianGames.xlsx", "PktReifenwuchten"))
PktBulldogziehen = DataFrame(XLSX.readtable("BavarianGames.xlsx", "PktBulldogziehen"))
PktBierkruglauf = DataFrame(XLSX.readtable("BavarianGames.xlsx", "PktBierkruglauf"))


FISPkt = [100 80 60 50 45 40 36 32 29 26 24 22 20 18 16 15 14 13 12 11 10 9 8 7 6 5 4 3 2 1]


begin # Masskrugschieben
  PktGesamtMKS = []
  for row in eachrow(PktMasskrugschieben)
    push!(PktGesamtMKS, sum(Vector(row[2:end])))
  end
  PktMasskrugschieben[!, :Gesamt] = PktGesamtMKS
  sort!(PktMasskrugschieben, [order(:Gesamt, rev=true)])
  println(PktMasskrugschieben)
end

begin # Wackelturm
  sort!(PktWackelturm, [order(:Hoehe, rev=true)])
  println(PktWackelturm)
end

begin # Holzscheitelwurf
  WeiteGesamtHSW = []
  for row in eachrow(PktHolzscheitelwurf)
    push!(WeiteGesamtHSW, sum(Vector(row[2:end])))
  end
  PktHolzscheitelwurf[!, :Gesamt] = WeiteGesamtHSW
  sort!(PktHolzscheitelwurf, [order(:Gesamt, rev=true)])
  println(PktHolzscheitelwurf)
end

begin # Reifenwuchten
  sort!(PktReifenwuchten, [order(:Zeit)])
  println(PktReifenwuchten)
end

begin # Bulldogziehen
  sort!(PktBulldogziehen, [order(:Zeit)])
  println(PktBulldogziehen)
end

begin # Bierkruglauf
  BierkruglaufGesamt = []
  for row in eachrow(PktBierkruglauf)
    push!(BierkruglaufGesamt, (row[2] - row[3]) / row[4])
  end
  PktBierkruglauf[!, :Gesamt] = BierkruglaufGesamt
  sort!(PktBierkruglauf, [order(:Gesamt, rev=true)])
  println(PktBierkruglauf)
end


begin # Gesamtauswertung
  Gruppen[!, :Gesamt] .= 0
  for i in 1:nrow(PktMasskrugschieben)
    GruppenID = PktMasskrugschieben[i, :GruppenID]
    Gruppen[GruppenID, :Gesamt] += FISPkt[i]
  end
  for i in 1:nrow(PktWackelturm)
    GruppenID = PktWackelturm[i, :GruppenID]
    Gruppen[GruppenID, :Gesamt] += FISPkt[i]
  end
  for i in 1:nrow(PktHolzscheitelwurf)
    GruppenID = PktHolzscheitelwurf[i, :GruppenID]
    Gruppen[GruppenID, :Gesamt] += FISPkt[i]
  end
  for i in 1:nrow(PktReifenwuchten)
    GruppenID = PktReifenwuchten[i, :GruppenID]
    Gruppen[GruppenID, :Gesamt] += FISPkt[i]
  end
  for i in 1:nrow(PktBulldogziehen)
    GruppenID = PktBulldogziehen[i, :GruppenID]
    Gruppen[GruppenID, :Gesamt] += FISPkt[i]
  end
  for i in 1:nrow(PktBierkruglauf)
    GruppenID = PktBierkruglauf[i, :GruppenID]
    Gruppen[GruppenID, :Gesamt] += FISPkt[i]
  end
  Bestenliste = sort(Gruppen, [order(:Gesamt, rev=true)])
  println(Bestenliste)
end
end
