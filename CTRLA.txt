import powerfactory as pf
app = pf.GetApplication()
app.ClearOutputWindow()

# Crear listas para clasificar
lineas = []
trafos2 = []
trafos3 = []
barras_visibles = []
shunts = []
switchs= []

# Obtener diagrama activo
diagram = app.GetCurrentDiagram()
if not diagram:
    app.PrintPlain("❌ No hay un diagrama activo.")
    quit()

# Obtener elementos del diagrama
selectedElements = diagram.GetContents(0)
if not selectedElements:
    app.PrintPlain("❌ No hay elementos visibles en el diagrama.")
    quit()

# Clasificar por tipo
for element in selectedElements:
    if element.loc_name in ["Settings", "Layers", "ColLeg"]:
        continue

    sym = element.GetAttribute("sSymNam")
    dataobject = element.GetAttribute("pDataObj")
    if not dataobject:
        continue

    if sym == "d_lin":
        lineas.append(dataobject)
    elif sym in ["d_tr2", "d_tr2ansi"]:
        trafos2.append(dataobject)
    elif sym == "d_tr3":
        trafos3.append(dataobject)
    elif sym == "d_shunt":
        shunts.append(dataobject)
    elif sym == "d_coupleRect":
        switchs.append(dataobject)
    elif sym in ["PointTerm", "RectTerm", "TermStrip", "ShortTermStrip"]:
        barras_visibles.append(dataobject)

# Imprimir en PowerFactory
app.PrintPlain(f"\n🔹 Líneas detectadas ({len(lineas)}):")
for obj in lineas:
    app.PrintPlain(f"Nombre: {obj.loc_name} Objeto: {str(obj)}")
    
app.PrintPlain(f"\n🔹 Trafo 2 devanados ({len(trafos2)}):")
for obj in trafos2:
    app.PrintPlain(f"- {obj.loc_name}")

app.PrintPlain(f"\n🔹 Trafo 3 devanados ({len(trafos3)}):")
for obj in trafos3:
    app.PrintPlain(f"- {obj.loc_name}")

app.PrintPlain(f"\n🔹 Barras visibles en el diagrama ({len(barras_visibles)}):")
for obj in barras_visibles:
    app.PrintPlain(f"- {obj.loc_name}")

app.PrintPlain(f"\n🔹 Switches visibles en el diagrama ({len(switchs)}):")
for obj in switchs:
    app.PrintPlain(f"- {obj.loc_name}")
    
app.PrintPlain(f"\n🔹 Capacitores visibles en el diagrama ({len(shunts)}):")
for obj in shunts:
    app.PrintPlain(f"- {obj.loc_name}")

app.PrintPlain(f"\n✅ Archivo generado con {len(lineas) + len(trafos2) + len(trafos3) + len(barras_visibles) + len(shunts) + len(switchs)} ")
