# Сборка: python3 src/build.py  → index.html (ПТО), director.html (руководитель), apps-script/Code.gs
# С ключами (только для себя, не в репозиторий): python3 src/build.py --keys ПТО_КЛЮЧ РУК_КЛЮЧ --out папка
import os, sys, json
d = os.path.dirname(os.path.abspath(__file__)); r = os.path.dirname(d)
rd = lambda f: open(os.path.join(d, f), encoding='utf-8').read()
css, pcss, logo, calc, ui = rd('base.css'), rd('pto.css'), rd('logo.txt').strip(), rd('calc.js'), rd('ui.js')
for tpl, out in [('app.tpl.html', 'index.html'), ('director.tpl.html', 'director.html')]:
    t = rd(tpl).replace('{{CSS}}', css).replace('{{PTOCSS}}', pcss).replace('{{LOGO}}', logo).replace('{{CALC}}', calc).replace('{{UI}}', ui)
    open(os.path.join(r, out), 'w', encoding='utf-8').write(t)
schema = json.dumps(json.load(open(os.path.join(d, 'schema.json'), encoding='utf-8')), ensure_ascii=False)
gs = open(os.path.join(r, 'apps-script', 'Code.tpl.gs'), encoding='utf-8').read().replace('{{SCHEMA}}', schema)
keys = ('ВСТАВЬТЕ_КЛЮЧ_ПТО', 'ВСТАВЬТЕ_КЛЮЧ_РУКОВОДИТЕЛЯ'); outdir = os.path.join(r, 'apps-script')
if '--keys' in sys.argv:
    i = sys.argv.index('--keys'); keys = (sys.argv[i + 1], sys.argv[i + 2]); outdir = sys.argv[sys.argv.index('--out') + 1]
open(os.path.join(outdir, 'Code.gs'), 'w', encoding='utf-8').write(gs.replace('{{PTO_KEY}}', keys[0]).replace('{{DIRECTOR_KEY}}', keys[1]))
print('ok')
