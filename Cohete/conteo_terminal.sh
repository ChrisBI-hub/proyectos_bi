#!/bin/bash

# 1. Enviar los correos de inmediato (4:20 PM)
echo "Enviando correos al equipo..."
/home/christian/Documentos/proyectos_bi/venv_bi/bin/python3 /home/christian/Documentos/proyectos_bi/enviar_artemis.py

# 2. Configurar el tiempo (4 minutos = 240 segundos)
segundos=240

echo "--- CONTEO REGRESIVO PARA EL LANZAMIENTO ARTEMIS 2 ---"

while [ $segundos -gt 0 ]; do
    # Calculamos minutos y segundos para mostrar
    minutos=$((segundos / 60))
    resto_seg=$((segundos % 60))
    
    # Imprimimos en la misma línea para que parezca un reloj
    printf "\rTIEMPO PARA EL DESPEGUE: %02d:%02d " $minutos $resto_seg
    
    sleep 1
    : $((segundos--))
done

echo -e "\n¡IGNICIÓN! Lanzando animación..."

# 3. Ejecutar el HTML a las 4:24 PM
firefox /home/christian/Documentos/proyectos_bi/cohete2.html &