from docxtpl import DocxTemplate
from datetime import datetime
import pandas

doc = DocxTemplate("wordTemplate.docx")

mi_empresa = "GMTECH SA"
mi_nombre = "Gaston Picon"
mi_dni = "26666673"
mi_fecha = datetime.today().strftime("%d %b, %Y")


mi_context = {'mi_empresa':mi_empresa, 'mi_nombre':mi_nombre,
              'mi_dni':mi_dni, 'mi_fecha':mi_fecha}

dataSource = pandas.read_csv('dataSource.csv')

for index, fila in dataSource.iterrows():
    context = {
        'var_empresa':fila['empresa'],
        'var_nombre':fila['nombre'],
        'var_dni':fila['dni']
    }

    context.update(mi_context)
#print(context)
    doc.render(context)
    doc.save(f"generated_doc_{fila['dni']}.docx")
#    print(montotito)
#    print(fila)
