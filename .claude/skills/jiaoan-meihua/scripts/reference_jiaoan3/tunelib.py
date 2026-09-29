import subprocess, os
BASE = dict(FILL_P1='0', FILL_P2='0', FILL_P3='0', FILL_P4='0', REFLECT_H='3000')
def pages(**kw):
    env = dict(os.environ, **BASE); env.update({k: str(v) for k, v in kw.items()})
    subprocess.run(['./build.sh', 't'], env=env, check=True, cwd='/tmp/lo1')
    x = subprocess.run(['pdftotext', 't.pdf', '-'], capture_output=True, text=True, cwd='/tmp/lo1').stdout
    return [pg.replace('\n', '').replace(' ', '') for pg in x.split('\f')][:-1]
def show(p):
    for i, pg in enumerate(p): print(i + 1, pg[:24], '...', pg[-24:])
