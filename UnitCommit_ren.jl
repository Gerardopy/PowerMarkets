# Unit Commitment Problem is based on Prof. Rodrigo Moreno's course on Reliability and Risk Analysis of Power Systems at the Universidad de Chile.
# https://www.notion.so/gerardoblanco/Unidad-VI-Operaci-n-del-Sistema-de-Potencia-d67ac0aa7da345e89048b13416f21df3
using JuMP, HiGHS, XLSX, DataFrames, Plots

function run()
    # Cargar datos desde un archivo Excel
# Cargar datos desde un archivo Excel (esto requerirá una conversión previa a CSV o similar)
    filename= "UnitCom1renovable.xlsx"
    xf = XLSX.readxlsx(filename)[2][:]
    df = DataFrame(xf[2:end,:],Symbol.(xf[1,:]), makeunique=true)
    Dem = df[!, "Dem Neta"]

    
    # Crear el modelo con el optimizador HiGHS
    model = Model(HiGHS.Optimizer)
    
    # Declaración de variables
    @variable(model, 0 <= Ph[1:24] <= 100) # Generador Hidroelêctrico
    @variable(model, 0 <= Pc[1:24] <= 500) # Generador a Carbón
    @variable(model, 0 <= Pg[1:24] <= 500) # Generador a Gas
    @variable(model, 0 <= Po[1:24] <= 500) # Generador Diesel
    @variable(model, 0 <= Pf[1:24] <= 9999) # Generador de Falla
    @variable(model, Xh[1:24], Bin) #Variable binaria de Estado de Generador Hidroelêctrico
    @variable(model, Xc[1:24], Bin) #Variable binaria de Estado de Generador Carbón
    @variable(model, Xg[1:24], Bin) #Variable binaria de Estado de Generador Gas
    @variable(model, Xo[1:24], Bin) #Variable binaria de Estado de Generador Diesel
    @variable(model, Eh[1:24], Bin) #Variable binaria de Encendido Generador Hidroelêctrico
    @variable(model, Ec[1:24], Bin) #Variable binaria de Encendido Generador Carbon
    @variable(model, Eg[1:24], Bin) #Variable binaria de Encendido Generador Gas
    @variable(model, Eo[1:24], Bin) #Variable binaria de Encendido Generador Diesel
    @variable(model, Ah[1:24], Bin) #Variable binaria de Apagado Generador Hidroelêctrico
    @variable(model, Ac[1:24], Bin) #Variable binaria de Apagado Generador Carbon
    @variable(model, Ag[1:24], Bin) #Variable binaria de Apagado Generador Gas
    @variable(model, Ao[1:24], Bin) #Variable binaria de Apagado Generador Diesel
    
    
    
    
    # Restricciones
    @constraint(model, balance[t in 1:24], Ph[t] + Pc[t] + Pg[t] + Po[t] + Pf[t]== Dem[t])
    @constraint(model, [t in 1:24], Ph[t] <= 100 * Xh[t])
    @constraint(model, [t in 1:24], Pc[t] <= 500 * Xc[t])
    @constraint(model, [t in 1:24], Pg[t] <= 500 * Xg[t])
    @constraint(model, [t in 1:24], Po[t] <= 500 * Xo[t])
    @constraint(model, [t in 1:24], Ph[t] >= 1 * Xh[t])
    @constraint(model, [t in 1:24], Pc[t] >= 30 * Xc[t])
    @constraint(model, [t in 1:24], Pg[t] >= 20 * Xg[t])
    @constraint(model, [t in 1:24], Po[t] >= 1 * Xo[t])
    @constraint(model, [t in 1], Xh[t] == Eh[t] - Ah[t])
    @constraint(model, [t in 1], Xc[t] == Ec[t] - Ac[t])
    @constraint(model, [t in 1], Xg[t] == Eg[t] - Ag[t])
    @constraint(model, [t in 1], Xo[t] == Eo[t] - Ao[t])
    @constraint(model, [t in 2:24], Xh[t] == Xh[t-1] + Eh[t] - Ah[t])
    @constraint(model, [t in 2:24], Xc[t] == Xc[t-1] + Ec[t] - Ac[t])
    @constraint(model, [t in 2:24], Xg[t] == Xg[t-1] + Eg[t] - Ag[t])
    @constraint(model, [t in 2:24], Xo[t] == Xo[t-1] + Eo[t] - Ao[t])
    @constraint(model, [t in 2:24], Ph[t] - Ph[t-1] <= Xh[t-1]*100 + Eh[t] * 100)
    @constraint(model, [t in 2:24], Pc[t] - Pc[t-1] <= Xc[t-1]*100 + Ec[t] * 500)
    @constraint(model, [t in 2:24], Pg[t] - Pg[t-1] <= Xg[t-1]*200 + Eg[t] * 500)
    @constraint(model, [t in 2:24], Po[t] - Po[t-1] <= (Xo[t-1] + Eo[t]) * 500)
    @constraint(model, [t in 2:24], Ph[t-1] - Ph[t] <= (Xh[t] + Ah[t]) * 100)
    @constraint(model, [t in 2:24], Pc[t-1] - Pc[t] <= (Xc[t] + Ac[t]) * 100)
    @constraint(model, [t in 2:24], Pg[t-1] - Pg[t] <= (Xg[t] + Ag[t]) * 200)
    @constraint(model, [t in 2:24], Po[t-1] - Po[t] <= (Xo[t] + Ao[t]) * 500)
    @constraint(model, [t in 4:24], Xh[t] >= sum(Eh[k] for k in (t-3):t))
    @constraint(model, [t in 4:24], Xc[t] >= sum(Ec[k] for k in (t-3):t))
    @constraint(model, [t in 4:24], Xg[t] >= sum(Eg[k] for k in (t-3):t))
   # @constraint(model, [t in 4:24], Xo[t] >= sum(Eo[k] for k in (t-3):t))
    @constraint(model, [t in 4:24], 1 - Xh[t] >= sum(Ah[k] for k in (t-3):t))
    @constraint(model, [t in 4:24], 1 - Xc[t] >= sum(Ac[k] for k in (t-3):t))
    @constraint(model, [t in 4:24], 1 - Xg[t] >= sum(Ag[k] for k in (t-3):t))
   # @constraint(model, [t in 4:24], 1 - Xo[t] >= sum(Ao[k] for k in (t-3):t))
    @constraint(model, Xh[1] >= Eh[1])
    @constraint(model, Xc[1] >= Ec[1])
    @constraint(model, Xg[1] >= Eg[1])
    #@constraint(model, Xo[1] >= Eo[1])
    @constraint(model, Xh[1] >= Ah[1])
    @constraint(model, Xc[1] >= Ac[1])
    @constraint(model, Xg[1] >= Ag[1])
    #@constraint(model, Xo[1] >= Ao[1])
    @constraint(model, Xh[2] >= Eh[2])
    @constraint(model, Xc[2] >= Ec[2])
    @constraint(model, Xg[2] >= Eg[2])
    #@constraint(model, Xo[2] >= Eo[2])
    @constraint(model, Xh[2] >= Ah[2])
    @constraint(model, Xc[2] >= Ac[2])
    @constraint(model, Xg[2] >= Ag[2])
    #@constraint(model, Xo[2] >= Ao[2])
    @constraint(model, Xh[3] >= Eh[2])
    @constraint(model, Xc[3] >= Ec[2])
    @constraint(model, Xg[3] >= Eg[3])
    #@constraint(model, Xo[3] >= Eo[3])
    @constraint(model, Xh[3] >= Ah[3])
    @constraint(model, Xc[3] >= Ac[3])
    @constraint(model, Xg[3] >= Ag[3])
    #@constraint(model, Xo[3] >= Ao[3])
    
    
# Función objetivo
@objective(model, Min, sum(0*Ph[t]+50 * Pc[t] + 100 * Pg[t] + 150 * Po[t] + 5000*Pf[t] + 5*Ec[t] + 10*Ec[t] + 15*Eo[t] + 6*Ao[t] + 11*Ag[t] +16*Ao[t] for t in 1:24))
    
    # Resolver el modelo
    optimize!(model)
    
    if termination_status(model) == MOI.OPTIMAL
        obj_value = objective_value(model)
        println("El valor objetivo es ", obj_value)
        foreach(t -> println("Hora $t: Pc=$(round(value(Pc[t]), digits=2)), Pg=$(round(value(Pg[t]), digits=2)), Po=$(round(value(Po[t]), digits=2)), Eh=$(value(Eh[t]))"), 1:24)
    else
        println("El modelo no tiene solución óptima o es infeasible. Estado: ", termination_status(model))
    end
    
    PM=zeros(24)

    XLSX.openxlsx(filename, mode="rw") do xg
    sheet = xg[2]
    for i in 1:24
        row_index=1+i
        sheet["D$(string(row_index))"] = value(Ph[i])
        sheet["E$(string(row_index))"] = value(Pc[i])
        sheet["F$(string(row_index))"] = value(Pg[i])
        sheet["G$(string(row_index))"] = value(Po[i])
        sheet["H$(string(row_index))"] = value(Pf[i])
        #sheet["M$(string(row_index))"] = dual(balance[i])
        
    end
end


run()
