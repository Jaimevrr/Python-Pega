import os
import shutil
import re

# Rutas
file_path = r'C:\Users\Jaime Valderrama\OneDrive - American Sportswear, S.A\Documentos\Jaime\Testing\Nueva carpeta\fotos unides CH4\CK 1'
output_path = r'C:\Users\Jaime Valderrama\OneDrive - American Sportswear, S.A\Documentos\Jaime\Testing\Nueva carpeta\Resultado\fotoresumen'

# Crear carpeta destino si no existe
os.makedirs(output_path, exist_ok=True)

# Diccionario para guardar la mejor imagen (la de menor número)
best_images = {}

# Expresión regular para extraer datos
pattern = re.compile(r'^([A-Z0-9]+)_([0-9]+)_([0-9]+)\.(jpg|jpeg|png)$', re.IGNORECASE)

# Recorremos las imágenes
for filename in os.listdir(file_path):
    match = pattern.match(filename)
    if match:
        referencia, color, numero_str, ext = match.groups()
        numero = int(numero_str)
        clave = f"{referencia}_{color}"
        
        # Guardamos solo la imagen de menor número por clave
        if clave not in best_images or numero < best_images[clave][1]:
            best_images[clave] = (filename, numero)

# Lista para Excel
referencias_excel = []

# Copiar archivos seleccionados
for clave, (filename, _) in best_images.items():
    src = os.path.join(file_path, filename)
    dst = os.path.join(output_path, filename)
    shutil.copy2(src, dst)
    referencias_excel.append(clave.replace("_", ""))  # sin guiones bajos

# Guardar lista como archivo txt (se puede abrir en Excel)
txt_output = os.path.join(output_path, "referencias.txt")
with open(txt_output, "w") as f:
    for ref in referencias_excel:
        f.write(ref + "\n")

print(f"Proceso completado. Se guardaron {len(referencias_excel)} imágenes y la lista en 'referencias.txt'.")
