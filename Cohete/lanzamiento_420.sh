#!/bin/bash

# 1. Abrir la animación en tu navegador local (Para que tú la veas)
# Usamos & para que se ejecute en segundo plano y el script continúe
firefox /home/christian/Documentos/proyectos_bi/cohete2.html &

# 2. Esperar 2 segundos para asegurar estabilidad
sleep 2

# 3. Ejecutar el script de Python para enviar los correos al equipo
# Usamos el python de tu entorno virtual si es necesario
/home/christian/Documentos/proyectos_bi/venv_bi/bin/python3 /home/christian/Documentos/proyectos_bi/enviar_artemis.py

echo "Sistemas iniciados: Animación abierta y correos enviados."