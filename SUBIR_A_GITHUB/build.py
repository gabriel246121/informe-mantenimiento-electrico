# Une el formulario, el generador, JSZip y el logo en un solo index.html
# Uso:  python build.py   (desde la carpeta del proyecto)
import base64, os
R = os.path.dirname(os.path.abspath(__file__))
src = lambda f: os.path.join(R, 'src', f)
s = open(src('index.src.html'), encoding='utf-8').read()
s = s.replace('/*JSZIP*/', open(src('jszip.min.js'), encoding='utf-8').read().replace('</script', '<\\/script'))
s = s.replace('/*GEN*/', open(src('gen.js'), encoding='utf-8').read())
s = s.replace('__LOGO__', base64.b64encode(open(src('logo.png'), 'rb').read()).decode())
open(os.path.join(R, 'index.html'), 'w', encoding='utf-8').write(s)
print('index.html generado:', len(s), 'caracteres')
print('Recuerda: si publicas cambios, sube la VERSION en sw.js (ej. informe-v2)')
