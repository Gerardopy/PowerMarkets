using JuMP,GLPK, XLSX, DataFrames, Plots

    
# Cargar datos desde un archivo Excel (esto requerirá una conversión previa a CSV o similar)
filename= "UnitCom1renovable.xlsx"
xf = XLSX.readxlsx(filename)[2][:]
df = DataFrame(xf[2:end,:],Symbol.(xf[1,:]), makeunique=true)
Dem = df[!, "Dem Neta"]

# Crear el modelo con el optimizador HiGHS
model = Model(GLPK.Optimizer)

# Declaración de variables
@variable(model, 0 <= Ph[1:24] <= 100) # Generador Hidroelêctrico
@variable(model, 0 <= Pc[1:24] <= 500) # Generador de Carbón
@variable(model, 0 <= Pg[1:24] <= 500) # Generador de Gas
@variable(model, 0 <= Po[1:24] <= 500) # Generador Diesel
@variable(model, 0 <= Pf[1:24] <= 9999) # Generador de Falla


# Restricciones
@constraint(model, balance[t in 1:24], Ph[t] + Pc[t] + Pg[t] + Po[t] + Pf[t]== Dem[t])

# Función objetivo
@objective(model, Min, sum(0*Ph[t] + 50*Pc[t] + 100*Pg[t] + 150*Po[t] + 5000*Pf[t] for t in 1:24))

# Resolver el modelo
optimize!(model)


if termination_status(model) == MOI.OPTIMAL
    obj_value = objective_value(model)
    println("El valor objetivo es ", obj_value)
    foreach(t -> println("Hora $t: Ph=$(round(value(Ph[t]), digits=2)), Pc=$(round(value(Pc[t]), digits=2)), Pg=$(round(value(Pg[t]), digits=2)), Po=$(round(value(Po[t]), digits=2)), Pf=$(round(value(Pf[t]), digits=2))"), 1:24)
    #foreach(t -> println("Hora $t: Precio marginal = $(Precio[t])"), 1:24)

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
        sheet["M$(string(row_index))"] = values(dual(balance[i]))
        
    end
end


for j in 1:24
    PM[j]=values(dual(balance[j]))
end
x=range(0,24,length=24)
precios = plot(x,PM,label="Precio Marginal", lw=:2, fmt=:png)
title!("Precios marginales horarios")
xlabel!("Horas")
ylabel!("USD/MWh")
savefig(precios,"Precios")
display(precios)

@userplot CirclePlot
@recipe function f(cp::CirclePlot)
    x, y, i = cp.args
    n = length(x)
    inds = circshift(1:n, 1 - i)
    linewidth --> range(0, 10, length = n)
    seriesalpha --> range(0, 1, length = n)
    aspect_ratio --> 0.1
    label --> false
    x[inds], y[inds]
end

n = 150
t = range(1, 0.5π, length = n)
y = PM

anim = @animate for i ∈ 1:n
    circleplot(x, y, i, linecolor=:red, ylims=(40, 160), xlims=(0,24))
end 
gif(anim, "anim_fps15.gif", fps = 60)
