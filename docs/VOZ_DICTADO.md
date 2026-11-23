# Dictado y lectura en voz alta

Fecha: **2026-11-22**  
Base: **Mini Word 0.0.5 Beta**

## Resumen

Sí es viable, pero son dos funciones distintas:

- **Leer en voz alta** es sencillo y ocupa muy poco.
- **Dictar** depende de que Windows tenga instalado un reconocedor de voz para el idioma.

La voz no debe ser obligatoria ni enviar audio a Internet por defecto.

## Prueba realizada en este equipo

Se compiló una prueba con `System.Speech`:

```text
SpeechSynthesizer: NullReferenceException
Reconocedores instalados: 0
```

Eso significa que **este Windows no tiene voces SAPI utilizables ni reconocedor instalado**. En un equipo con los paquetes de voz de Windows el resultado puede ser distinto. La aplicación debe detectar esa ausencia y mostrar un aviso, nunca bloquearse.

## Leer en voz alta

### Tecnología recomendada

`System.Speech.Synthesis.SpeechSynthesizer`, incluido en .NET Framework.

Ventajas:

- No requiere instalar nada adicional si Windows tiene voces SAPI.
- No usa Internet.
- Permite pausar, reanudar, detener y elegir velocidad.
- `SpeakProgress` permite resaltar la palabra o frase que se está leyendo.

### Implementación

1. Añadir referencia a `System.Speech.dll`.
2. Crear un `ReadAloudService`:
   - `CanSpeak`.
   - `Voices`.
   - `Speak(text, start, length)`.
   - `Pause`, `Resume`, `Stop`.
   - Evento `Progress` con posición de carácter.
3. Convertir el documento o la selección a texto plano con un mapa de posiciones:
   - Párrafo.
   - Offset inicial.
   - Longitud.
4. Al recibir `SpeakProgress`, seleccionar el rango correspondiente.
5. Añadir **Revisar → Leer en voz alta** y **Detener**.
6. Si no hay voces:
   - Mostrar “Windows no tiene una voz instalada”.
   - Enlace a Configuración → Voz.

### Espacio adicional

- Código: aproximadamente **20–40 KB**.
- `System.Speech.dll`: ya forma parte de .NET Framework; no se empaqueta.
- Voces SAPI: las instala Windows; pueden ocupar **50–200 MB** por idioma, fuera del ejecutable.
- No requiere modelos de IA.

## Dictado

### Tecnología local

`System.Speech.Recognition.SpeechRecognitionEngine` con `DictationGrammar`.

Requisitos:

- Micrófono.
- Permiso de micrófono.
- Reconocedor de voz instalado para el idioma.
- Paquete de idioma de Windows.

En la prueba de este equipo:

```text
InstalledRecognizers = 0
```

Sin un reconocedor, el dictado no puede funcionar. No es un fallo de Mini Word: falta el componente de Windows.

### Implementación

1. Crear `SpeechRecognitionEngine` con la cultura configurada.
2. Cargar `DictationGrammar`.
3. `SetInputToDefaultAudioDevice()`.
4. `RecognizeAsync(RecognizeMode.Multiple)`.
5. En `SpeechRecognized`, insertar el texto en el cursor.
6. Detener y liberar al cerrar.
7. Manejar:
   - Sin micrófono.
   - Sin reconocedor.
   - Permiso denegado.
   - Idioma no disponible.
8. Comandos opcionales:
   - “punto”, “coma”, “nueva línea”.
   - “borrar”.
   - “detener dictado”.

### Calidad

El reconocimiento SAPI de Windows es aceptable para frases claras, pero inferior a Whisper o a los servicios en la nube. En español depende mucho del paquete instalado.

## Alternativas

### Windows.Media.SpeechSynthesis / SpeechRecognition

API moderna de Windows 10/11.

- Mejores voces OneCore.
- Reconocimiento integrado.
- Requiere interoperabilidad WinRT desde .NET Framework.
- Más complejo que `System.Speech`.
- No añade modelos al ejecutable.
- Depende de los paquetes de idioma de Windows.

### Whisper local

- Calidad alta.
- Modelos:
  - `tiny`: ~75 MB.
  - `base`: ~150 MB.
  - `small`: ~500 MB.
  - `medium`: ~1,5 GB.
- Runtime nativo adicional.
- Consume CPU/RAM.
- No recomendado para una aplicación de 439 KB.

### Servicios en la nube

- Azure Speech, OpenAI Whisper, etc.
- No ocupan disco.
- Requieren clave, Internet y consentimiento.
- El audio sale del equipo.
- No debería ser la opción predeterminada.

## Espacio estimado

| Opción | Espacio en la app | Espacio externo |
|---|---:|---:|
| Leer con SAPI | 20–40 KB | Voces de Windows: 50–200 MB |
| Dictar con SAPI | 40–80 KB | Paquete de voz: 50–200 MB |
| Windows.Media | 50–100 KB | Paquetes de Windows |
| Whisper tiny local | 100–300 KB de código | ~75 MB de modelo + runtime |
| Whisper small local | 100–300 KB | ~500 MB + runtime |
| Nube | 20–50 KB | 0 |

## Privacidad y seguridad

Recomendación:

- Voz **local por defecto**.
- Micrófono solo mientras se pulsa **Dictar**.
- Indicador visible de grabación.
- No guardar audio.
- No enviar texto ni audio a Internet sin consentimiento explícito.
- Detener el reconocimiento al cerrar el documento.
- Liberar micrófono y motor al terminar.

## Plan de implementación

### Fase 1: leer en voz alta

1. Referencia `System.Speech`.
2. Servicio de síntesis.
3. Botones Leer/Detener.
4. Resaltado por `SpeakProgress`.
5. Manejo de ausencia de voces.

**Esfuerzo:** bajo-medio.  
**Riesgo:** bajo.

### Fase 2: dictado local

1. Detectar reconocedores.
2. Motor de dictado.
3. Inserción en el cursor.
4. Comandos de puntuación.
5. Manejo de permisos y errores.
6. Pruebas en un equipo con paquete de voz español.

**Esfuerzo:** medio.  
**Riesgo:** medio por disponibilidad del idioma y micrófono.

### Fase 3: voces modernas opcionales

- Windows.Media para mejores voces.
- Proveedor en la nube opcional.
- Whisper local como descarga separada.

## Conclusión

**Leer en voz alta es muy viable** y añade muy poco tamaño.

**Dictar es viable solo si Windows tiene un reconocedor del idioma.** En este equipo no lo hay. La implementación debe ser opcional, detectar la ausencia y explicar cómo instalar el paquete de voz.

No conviene empaquetar Whisper en Mini Word: multiplicaría por decenas o cientos el tamaño del ejecutable.
